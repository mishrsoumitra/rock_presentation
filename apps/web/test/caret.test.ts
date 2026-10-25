// Caret movement and hit testing over real laid-out text, in the file's offsets.
import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, test } from 'vitest';
import { OpcPackage } from '@pptx/opc';
import { PresentationDocument } from '@pptx/model';
import { shapeTextLayout } from '@pptx/render';
import { resolveSlide } from '@pptx/resolve';
import { FontRegistry } from '@pptx/text';
import { deckBytes } from '../../../packages/resolve/test/deck.ts';
import { caretRect, hit, move, wordAt, type TextView } from '../src/ui/caret.ts';

const fonts = new FontRegistry();
const FONTS = join(import.meta.dirname, '../../../assets/fonts');
for (const dir of ['carlito', 'caladea', 'liberation']) {
  for (const f of readdirSync(join(FONTS, dir))) if (f.endsWith('.ttf')) fonts.addBytes(new Uint8Array(readFileSync(join(FONTS, dir, f))));
}

/** The content placeholder of the test deck: "Revenue up" / "in every region" (level 2) / "details". */
function view(): TextView {
  const doc = new PresentationDocument(OpcPackage.open(deckBytes()));
  const s = resolveSlide(doc, '/ppt/slides/slide1.xml').shapes.find(x => x.name === 'Content Placeholder 2' && x.layer === 'slide')!;
  const t = shapeTextLayout(s, fonts)!;
  return { txBody: s.view.txBody!, resolved: s.text, layout: t.layout };
}

describe('caret movement', () => {
  const v = view();
  test('left and right cross paragraph ends', () => {
    expect(move(v, { p: 0, offset: 10 }, 'right').pos).toEqual({ p: 1, offset: 0 });
    expect(move(v, { p: 1, offset: 0 }, 'left').pos).toEqual({ p: 0, offset: 10 });
    expect(move(v, { p: 0, offset: 0 }, 'left').pos).toEqual({ p: 0, offset: 0 });
  });

  test('words: to the start of the word, then the previous one', () => {
    expect(move(v, { p: 1, offset: 15 }, 'wordLeft').pos).toEqual({ p: 1, offset: 9 });
    expect(move(v, { p: 1, offset: 0 }, 'wordRight').pos).toEqual({ p: 1, offset: 3 });
    expect(wordAt(v, { p: 1, offset: 5 })).toEqual({ start: { p: 1, offset: 3 }, end: { p: 1, offset: 8 } });
  });

  test('up and down keep the caret near its x; home and end go to the line ends', () => {
    const down = move(v, { p: 0, offset: 3 }, 'down');
    expect(down.pos.p).toBe(1);
    const back = move(v, down.pos, 'up', down.goalX);
    expect(back.pos).toEqual({ p: 0, offset: 3 });
    expect(move(v, { p: 0, offset: 3 }, 'end').pos).toEqual({ p: 0, offset: 10 });
    expect(move(v, { p: 0, offset: 3 }, 'home').pos).toEqual({ p: 0, offset: 0 });
  });

  test('a point gives the position whose caret is nearest it', () => {
    for (const pos of [{ p: 0, offset: 4 }, { p: 1, offset: 7 }, { p: 2, offset: 2 }]) {
      const r = caretRect(v, pos)!;
      expect(hit(v, r.x + 1, r.top + r.height / 2)).toEqual(pos);
    }
  });
});
