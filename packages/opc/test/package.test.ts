import { strToU8, unzipSync } from 'fflate';
import { describe, expect, test } from 'vitest';
import { canonical, childElements, getAttr, parseXml } from '@pptx/xml';
import { OpcError, OpcPackage, type OpcErrorCode } from '../src/index.ts';
import { deck, deckFiles, fakeCompoundFile, PNG_BYTES, REL_TYPE, zipFiles } from './fixtures.ts';

const dec = new TextDecoder();
const text = (b: Uint8Array | undefined) => dec.decode(b);

function fails(code: OpcErrorCode, fn: () => unknown): OpcError {
  let error: unknown;
  try {
    fn();
  } catch (e) {
    error = e;
  }
  expect(error).toBeInstanceOf(OpcError);
  expect((error as OpcError).code).toBe(code);
  return error as OpcError;
}

describe('opening', () => {
  test('lists parts in ZIP order and resolves content types: Override first, then Default by extension', () => {
    const pkg = OpcPackage.open(deck());
    expect(pkg.partNames()).toEqual([
      '/_rels/.rels', '/ppt/presentation.xml', '/ppt/_rels/presentation.xml.rels',
      '/ppt/slides/slide1.xml', '/ppt/slides/_rels/slide1.xml.rels', '/ppt/media/image1.png',
    ]);
    expect(pkg.getPart('/ppt/slides/slide1.xml')!.contentType).toBe('application/vnd.openxmlformats-officedocument.presentationml.slide+xml');
    expect(pkg.getPart('/ppt/media/image1.png')!.contentType).toBe('image/png');
    expect(pkg.getPart('/_rels/.rels')!.contentType).toBe('application/vnd.openxmlformats-package.relationships+xml');
    expect(pkg.diagnostics).toEqual([]);
  });

  test('part names match case-insensitively and through percent-escapes', () => {
    const files = deckFiles();
    files['ppt/media/my image.png'] = PNG_BYTES;
    const pkg = OpcPackage.open(zipFiles(files));
    expect(pkg.getPart('/PPT/Slides/SLIDE1.xml')?.name).toBe('/ppt/slides/slide1.xml');
    expect(pkg.getPart('ppt/slides/slide1.xml')?.name).toBe('/ppt/slides/slide1.xml');
    expect(pkg.getPart('/ppt/media/my%20image.png')?.name).toBe('/ppt/media/my image.png');
    expect(pkg.getPart('/ppt/media/%E0%A4%A')).toBeUndefined();   // a malformed escape is just a name
  });

  test('a part with no content type is kept and reported, not refused', () => {
    const files = deckFiles();
    files['ppt/media/clip.avi'] = new Uint8Array([1, 2, 3]);
    const pkg = OpcPackage.open(zipFiles(files));
    expect(pkg.getPart('/ppt/media/clip.avi')!.contentType).toBe('');
    expect(pkg.diagnostics).toEqual(['no content type for /ppt/media/clip.avi']);
    expect(unzipSync(pkg.save())['ppt/media/clip.avi']).toEqual(new Uint8Array([1, 2, 3]));
  });

  test('folder entries are skipped', () => {
    const files = deckFiles();
    const pkg = OpcPackage.open(zipFiles({ ...files, 'ppt/empty/': new Uint8Array(0) }));
    expect(pkg.partNames()).not.toContain('/ppt/empty/');
  });
});

describe('refusing what is not an Open XML package', () => {
  test('arbitrary bytes', () => {
    fails('not-a-package', () => OpcPackage.open(strToU8('%PDF-1.7 not a deck')));
    fails('not-a-package', () => OpcPackage.open(new Uint8Array(0)));
  });

  test('a ZIP without [Content_Types].xml', () => {
    const files = deckFiles();
    delete files['[Content_Types].xml'];
    fails('invalid-package', () => OpcPackage.open(zipFiles(files)));
  });

  test('a [Content_Types].xml that is not well-formed or has the wrong root', () => {
    fails('invalid-package', () => OpcPackage.open(zipFiles({ ...deckFiles(), '[Content_Types].xml': '<Types>' })));
    fails('invalid-package', () => OpcPackage.open(zipFiles({ ...deckFiles(), '[Content_Types].xml': '<Other/>' })));
  });

  test('ZIP entry names that cannot be part names', () => {
    for (const bad of ['../evil.xml', 'ppt/../../evil.xml', 'ppt/./slide.xml', 'ppt\\slide.xml', 'ppt/sl\u0001ide.xml']) {
      fails('unsafe-name', () => OpcPackage.open(zipFiles({ ...deckFiles(), [bad]: '<x/>' })));
    }
  });

  test('two ZIP entries that name the same part', () => {
    fails('invalid-package', () => OpcPackage.open(zipFiles({ ...deckFiles(), 'ppt/slides/SLIDE1.xml': '<x/>' })));
  });

  test('password-protected packages, legacy .ppt and other compound files get a clear reason', () => {
    expect(fails('encrypted', () => OpcPackage.open(fakeCompoundFile(['Root Entry', 'EncryptionInfo', 'EncryptedPackage']))).message)
      .toContain('password-protected');
    expect(fails('legacy-ppt', () => OpcPackage.open(fakeCompoundFile(['Root Entry', 'Current User', 'PowerPoint Document']))).message)
      .toContain('.ppt');
    fails('legacy-ppt', () => OpcPackage.open(fakeCompoundFile(['Root Entry', 'PP40'])));
    fails('unsupported-container', () => OpcPackage.open(fakeCompoundFile(['Root Entry', 'WordDocument'])));
  });
});

describe('saving', () => {
  test('with no edits, every entry comes back with identical bytes', () => {
    const original = deck();
    const a = unzipSync(original);
    const b = unzipSync(OpcPackage.open(original).save());
    expect(Object.keys(b)).toEqual(Object.keys(a));
    for (const name of Object.keys(a)) expect(b[name], name).toEqual(a[name]);
  });

  test('saving is deterministic and idempotent', () => {
    const once = OpcPackage.open(deck()).save();
    expect(OpcPackage.open(deck()).save()).toEqual(once);
    expect(OpcPackage.open(once).save()).toEqual(once);
  });

  test('a parsed but unedited part is written from its original bytes', () => {
    const pkg = OpcPackage.open(deck());
    pkg.xml('/ppt/slides/slide1.xml');
    expect(text(unzipSync(pkg.save())['ppt/slides/slide1.xml'])).toBe(deckFiles()['ppt/slides/slide1.xml']);
  });

  test('an edited part is re-serialized; everything else keeps its bytes', () => {
    const pkg = OpcPackage.open(deck());
    const doc = pkg.xml('/ppt/presentation.xml');
    const sldSz = childElements(doc.root).find(c => c.local === 'sldSz')!;
    sldSz.attrs.find(a => a.local === 'cx')!.value = '9144000';
    pkg.markDirty('/ppt/presentation.xml');
    const out = unzipSync(pkg.save());
    expect(text(out['ppt/presentation.xml'])).toBe((deckFiles()['ppt/presentation.xml'] as string).replace('12192000', '9144000'));
    expect(text(out['ppt/slides/slide1.xml'])).toBe(deckFiles()['ppt/slides/slide1.xml']);
    expect(text(out['[Content_Types].xml'])).toBe(deckFiles()['[Content_Types].xml']);
  });

  test('xml() returns the same tree each time', () => {
    const pkg = OpcPackage.open(deck());
    expect(pkg.xml('/ppt/slides/slide1.xml')).toBe(pkg.xml('/PPT/slides/slide1.xml'));
  });

  test('a malformed XML part fails only when parsed, and is still saved untouched', () => {
    const pkg = OpcPackage.open(zipFiles({ ...deckFiles(), 'ppt/slides/slide1.xml': '<p:sld><unclosed>' }));
    fails('invalid-part', () => pkg.xml('/ppt/slides/slide1.xml'));
    expect(text(unzipSync(pkg.save())['ppt/slides/slide1.xml'])).toBe('<p:sld><unclosed>');
  });

  test('a UTF-16 part is parsed and written back as UTF-16 with its BOM', () => {
    const xml = '<?xml version="1.0" encoding="UTF-16" standalone="yes"?>\r\n<r xmlns="urn:x">ünïcode</r>';
    const utf16 = new Uint8Array(2 + xml.length * 2);
    utf16.set([0xff, 0xfe]);
    for (let i = 0; i < xml.length; i++) { utf16[2 + 2 * i] = xml.charCodeAt(i) & 0xff; utf16[3 + 2 * i] = xml.charCodeAt(i) >> 8; }
    const pkg = OpcPackage.open(zipFiles({ ...deckFiles(), 'ppt/custom.xml': utf16 }));
    expect(canonical(pkg.xml('/ppt/custom.xml'))).toBe(canonical(parseXml(xml)));
    pkg.markDirty('/ppt/custom.xml');
    expect(unzipSync(pkg.save())['ppt/custom.xml']).toEqual(utf16);
  });

  test('setPartBytes replaces binary content', () => {
    const pkg = OpcPackage.open(deck());
    pkg.setPartBytes('/ppt/media/image1.png', new Uint8Array([9, 9, 9]));
    expect(unzipSync(pkg.save())['ppt/media/image1.png']).toEqual(new Uint8Array([9, 9, 9]));
  });
});

describe('relationships', () => {
  test('lists relationships of a part and of the package', () => {
    const pkg = OpcPackage.open(deck());
    expect(pkg.relationships('/')).toEqual([{ id: 'rId1', type: REL_TYPE.officeDocument, target: 'ppt/presentation.xml', external: false }]);
    expect(pkg.relationships('/ppt/slides/slide1.xml')).toEqual([
      { id: 'rId2', type: REL_TYPE.image, target: '../media/image1.png', external: false },
      { id: 'rId3', type: REL_TYPE.hyperlink, target: 'https://example.com/a?b=1&c=2', external: true },
    ]);
    expect(pkg.relationships('/ppt/media/image1.png')).toEqual([]);
  });

  test('follows relationships to parts', () => {
    const pkg = OpcPackage.open(deck());
    const presentation = pkg.related('/', 'rId1');
    expect(presentation.name).toBe('/ppt/presentation.xml');
    const slide = pkg.related(presentation.name, 'rId2');
    expect(pkg.related(slide.name, 'rId2').name).toBe('/ppt/media/image1.png');
  });

  test('missing, external and dangling relationships are errors', () => {
    const pkg = OpcPackage.open(deck());
    fails('missing-relationship', () => pkg.related('/ppt/slides/slide1.xml', 'rId99'));
    fails('missing-relationship', () => pkg.related('/ppt/slides/slide1.xml', 'rId3'));
    const files = deckFiles();
    delete files['ppt/media/image1.png'];
    fails('dangling-relationship', () => OpcPackage.open(zipFiles(files)).related('/ppt/slides/slide1.xml', 'rId2'));
  });

  test('addRelationship mints the next rId and writes a relative target', () => {
    const pkg = OpcPackage.open(deck());
    pkg.addPart('/ppt/media/image2.png', 'image/png', PNG_BYTES);
    const id = pkg.addRelationship('/ppt/slides/slide1.xml', REL_TYPE.image, '/ppt/media/image2.png');
    expect(id).toBe('rId4');
    expect(pkg.related('/ppt/slides/slide1.xml', id).name).toBe('/ppt/media/image2.png');
    const reopened = OpcPackage.open(pkg.save());
    expect(reopened.relationships('/ppt/slides/slide1.xml').at(-1)).toEqual({ id: 'rId4', type: REL_TYPE.image, target: '../media/image2.png', external: false });
  });

  test('addRelationship creates the relationships part when a part has none', () => {
    const pkg = OpcPackage.open(deck());
    pkg.addPart('/ppt/notesSlides/notesSlide1.xml', 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml', strToU8('<p:notes/>'));
    const id = pkg.addRelationship('/ppt/notesSlides/notesSlide1.xml', REL_TYPE.slide, '/ppt/slides/slide1.xml');
    expect(id).toBe('rId1');
    const out = unzipSync(pkg.save());
    expect(text(out['ppt/notesSlides/_rels/notesSlide1.xml.rels'])).toBe(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${REL_TYPE.slide}" Target="../slides/slide1.xml"/></Relationships>`);
  });

  test('external relationships keep their URI and TargetMode', () => {
    const pkg = OpcPackage.open(deck());
    const id = pkg.addRelationship('/ppt/slides/slide1.xml', REL_TYPE.hyperlink, 'https://example.org/x', { external: true });
    expect(pkg.relationships('/ppt/slides/slide1.xml').find(r => r.id === id)).toEqual({ id, type: REL_TYPE.hyperlink, target: 'https://example.org/x', external: true });
  });

  test('an internal relationship needs an existing target', () => {
    fails('missing-part', () => OpcPackage.open(deck()).addRelationship('/ppt/slides/slide1.xml', REL_TYPE.image, '/ppt/media/nope.png'));
  });

  test('removeRelationship and incomingRelationships', () => {
    const pkg = OpcPackage.open(deck());
    expect(pkg.incomingRelationships('/ppt/media/image1.png')).toEqual([{ source: '/ppt/slides/slide1.xml', id: 'rId2', type: REL_TYPE.image }]);
    pkg.removeRelationship('/ppt/slides/slide1.xml', 'rId2');
    expect(pkg.incomingRelationships('/ppt/media/image1.png')).toEqual([]);
    fails('missing-relationship', () => pkg.removeRelationship('/ppt/slides/slide1.xml', 'rId2'));
  });
});

describe('adding and removing parts', () => {
  test('a part whose extension has the right Default needs no Override, and [Content_Types].xml stays untouched', () => {
    const pkg = OpcPackage.open(deck());
    pkg.addPart('/ppt/media/image2.png', 'image/png', PNG_BYTES);
    const out = unzipSync(pkg.save());
    expect(text(out['[Content_Types].xml'])).toBe(deckFiles()['[Content_Types].xml']);
    expect(Object.keys(out).at(-1)).toBe('ppt/media/image2.png');
  });

  test('a part that needs an Override gets one, after the existing ones', () => {
    const pkg = OpcPackage.open(deck());
    pkg.addPart('/ppt/slides/slide2.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml', strToU8('<p:sld/>'));
    const types = parseXml(text(unzipSync(pkg.save())['[Content_Types].xml'])).root;
    const last = childElements(types).at(-1)!;
    expect([last.local, getAttr(last, 'PartName'), getAttr(last, 'ContentType')]).toEqual(
      ['Override', '/ppt/slides/slide2.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml']);
    expect(OpcPackage.open(pkg.save()).getPart('/ppt/slides/slide2.xml')!.contentType).toBe('application/vnd.openxmlformats-officedocument.presentationml.slide+xml');
  });

  test('adding an existing part is an error', () => {
    fails('part-exists', () => OpcPackage.open(deck()).addPart('/PPT/media/image1.png', 'image/png', PNG_BYTES));
    fails('unsafe-name', () => OpcPackage.open(deck()).addPart('/ppt/../x.png', 'image/png', PNG_BYTES));
  });

  test('removing a part that is still referenced is refused', () => {
    const pkg = OpcPackage.open(deck());
    expect(fails('referenced-part', () => pkg.removePart('/ppt/media/image1.png')).message).toContain('/ppt/slides/slide1.xml rId2');
  });

  test('removing a part takes its Override and its relationships part with it', () => {
    const pkg = OpcPackage.open(deck());
    pkg.removeRelationship('/ppt/presentation.xml', 'rId2');
    pkg.removePart('/ppt/slides/slide1.xml');
    const out = unzipSync(pkg.save());
    expect(Object.keys(out)).not.toContain('ppt/slides/slide1.xml');
    expect(Object.keys(out)).not.toContain('ppt/slides/_rels/slide1.xml.rels');
    expect(text(out['[Content_Types].xml'])).not.toContain('/ppt/slides/slide1.xml');
    expect(text(out['[Content_Types].xml'])).toContain('/ppt/presentation.xml');
    // the image is now unreferenced and can go
    pkg.removePart('/ppt/media/image1.png');
    expect(OpcPackage.open(pkg.save()).partNames()).toEqual(['/_rels/.rels', '/ppt/presentation.xml', '/ppt/_rels/presentation.xml.rels']);
  });

  test('force removes a referenced part anyway', () => {
    const pkg = OpcPackage.open(deck());
    pkg.removePart('/ppt/media/image1.png', { force: true });
    fails('dangling-relationship', () => pkg.related('/ppt/slides/slide1.xml', 'rId2'));
  });

  test('a new package built from scratch round-trips', () => {
    const pkg = OpcPackage.create();
    pkg.addPart('/ppt/presentation.xml', 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml', strToU8('<p:presentation xmlns:p="urn:p"/>'));
    pkg.addRelationship('/', REL_TYPE.officeDocument, '/ppt/presentation.xml');
    const reopened = OpcPackage.open(pkg.save());
    expect(reopened.related('/', 'rId1').name).toBe('/ppt/presentation.xml');
    expect(Object.keys(unzipSync(pkg.save()))).toEqual(['[Content_Types].xml', 'ppt/presentation.xml', '_rels/.rels']);
    expect(reopened.diagnostics).toEqual([]);
  });
});
