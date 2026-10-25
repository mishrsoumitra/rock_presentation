import { describe, expect, test } from 'vitest';
import { extensionOf, partKey, relativeTarget, relsNameFor, resolveTarget, sourceOfRels, unsafeEntryName } from '../src/index.ts';

describe('part names', () => {
  test('relationships part names', () => {
    expect(relsNameFor('/')).toBe('/_rels/.rels');
    expect(relsNameFor('/ppt/presentation.xml')).toBe('/ppt/_rels/presentation.xml.rels');
    expect(relsNameFor('/ppt/slides/slide1.xml')).toBe('/ppt/slides/_rels/slide1.xml.rels');
    expect(sourceOfRels('/_rels/.rels')).toBe('/');
    expect(sourceOfRels('/ppt/slides/_rels/slide1.xml.rels')).toBe('/ppt/slides/slide1.xml');
    expect(sourceOfRels('_rels/.rels')).toBe('/');
    expect(sourceOfRels('/ppt/_rels/.rels')).toBeNull();
    expect(sourceOfRels('/ppt/slides/slide1.xml')).toBeNull();
  });

  test('equivalence is case-insensitive ASCII after percent-decoding', () => {
    expect(partKey('/PPT/Slides/Slide1.XML')).toBe(partKey('ppt/slides/slide1.xml'));
    expect(partKey('/ppt/media/a%20b.png')).toBe(partKey('/ppt/media/a b.png'));
    expect(partKey('/ppt/Ä.xml')).not.toBe(partKey('/ppt/ä.xml'));   // only ASCII folds
    expect(partKey('/ppt/100%.xml')).toBe('/ppt/100%.xml');           // malformed escape: compared as written
  });

  test('extensions', () => {
    expect(extensionOf('/ppt/media/image1.PNG')).toBe('png');
    expect(extensionOf('/ppt/media/noext')).toBe('');
    expect(extensionOf('/ppt.d/noext')).toBe('');
    expect(extensionOf('/_rels/.rels')).toBe('rels');
  });

  test('unsafe ZIP entry names', () => {
    expect(unsafeEntryName('ppt/slides/slide1.xml')).toBeNull();
    expect(unsafeEntryName('[trash]/0000.dat')).toBeNull();
    expect(unsafeEntryName('../x')).toBe('dot segment in name');
    expect(unsafeEntryName('a/./b')).toBe('dot segment in name');
    expect(unsafeEntryName('/abs')).toBe('absolute path');
    expect(unsafeEntryName('a\\b')).toBe('backslash in name');
    expect(unsafeEntryName('a\nb')).toBe('control character in name');
    expect(unsafeEntryName('')).toBe('empty name');
    expect(unsafeEntryName('image..png')).toBeNull();                  // two dots inside a name are fine
  });
});

describe('relationship targets', () => {
  test('resolve against the source part', () => {
    expect(resolveTarget('/', 'ppt/presentation.xml')).toBe('/ppt/presentation.xml');
    expect(resolveTarget('/ppt/presentation.xml', 'slides/slide1.xml')).toBe('/ppt/slides/slide1.xml');
    expect(resolveTarget('/ppt/slides/slide1.xml', '../media/image1.png')).toBe('/ppt/media/image1.png');
    expect(resolveTarget('/ppt/slides/slide1.xml', './slide2.xml')).toBe('/ppt/slides/slide2.xml');
    expect(resolveTarget('/ppt/slides/slide1.xml', '/ppt/media/image1.png')).toBe('/ppt/media/image1.png');
    expect(resolveTarget('/ppt/slides/slide1.xml', '../../../../x.xml')).toBe('/x.xml');
    expect(resolveTarget('/ppt/slides/slide1.xml', 'slide2.xml#frag')).toBe('/ppt/slides/slide2.xml');
  });

  test('relative references are written the way PowerPoint writes them', () => {
    expect(relativeTarget('/', '/ppt/presentation.xml')).toBe('ppt/presentation.xml');
    expect(relativeTarget('/ppt/presentation.xml', '/ppt/slides/slide1.xml')).toBe('slides/slide1.xml');
    expect(relativeTarget('/ppt/slides/slide1.xml', '/ppt/media/image1.png')).toBe('../media/image1.png');
    expect(relativeTarget('/ppt/slides/slide1.xml', '/ppt/slides/slide2.xml')).toBe('slide2.xml');
    expect(relativeTarget('/ppt/slides/slide1.xml', '/docProps/thumbnail.jpeg')).toBe('../../docProps/thumbnail.jpeg');
  });

  test('relative and resolve are inverses', () => {
    const parts = ['/ppt/presentation.xml', '/ppt/slides/slide1.xml', '/ppt/media/image1.png', '/docProps/core.xml', '/customXml/item1.xml'];
    for (const from of ['/', ...parts]) for (const to of parts) expect(resolveTarget(from, relativeTarget(from, to))).toBe(to);
  });
});
