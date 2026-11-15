// Animation triggers as an undoable command: an effect started by a click on a shape, and back to the slide's clicks.
import { findPreset, listEffects } from '@pptx/anim';
import { describe, expect, test } from 'vitest';
import { deckBytes } from '../../resolve/test/deck.ts';
import { addAnimation, deleteShapes, insertShape, setAnimationTrigger, setShapeAnimationTrigger, type Editor } from '../src/index.ts';
import { cycle, openEditor } from './helpers.ts';

const SLIDE = '/ppt/slides/slide1.xml';
const fade = findPreset('entr', 10)!;
const effects = (ed: Editor) => listEffects(ed.slide(SLIDE)).map(e => [e.spid, e.triggerShape]);

function animated(): { ed: Editor; box: number; button: number } {
  const ed = openEditor(deckBytes());
  const a = insertShape(ed, SLIDE, 'rect', { x: 1000000, y: 3000000, w: 1000000, h: 1000000 });
  ed.run(a.command);
  const b = insertShape(ed, SLIDE, 'ellipse', { x: 4000000, y: 3000000, w: 1000000, h: 1000000 });
  ed.run(b.command);
  ed.run(addAnimation(ed, SLIDE, ed.shape(SLIDE, a.id)!, fade));
  ed.log.clear();
  return { ed, box: a.id, button: b.id };
}

describe('setAnimationTrigger', () => {
  test('an effect is given a shape to start it, and given back to the slide\'s clicks; each undone exactly', () => {
    const { ed, box, button } = animated();
    expect(effects(ed)).toEqual([[box, undefined]]);
    expect(cycle(ed, () => setAnimationTrigger(ed, SLIDE, 0, button))).toEqual([]);
    ed.run(setAnimationTrigger(ed, SLIDE, 0, button));
    expect(effects(ed)).toEqual([[box, button]]);
    expect(ed.log.undoLabel).toBe('Animation Trigger');
    expect(cycle(ed, () => setAnimationTrigger(ed, SLIDE, 0, null))).toEqual([]);
    ed.run(setAnimationTrigger(ed, SLIDE, 0, null));
    expect(effects(ed)).toEqual([[box, undefined]]);
    ed.undo();
    expect(effects(ed)).toEqual([[box, button]]);
  });

  test('a saved deck opens with the trigger; deleting the trigger shape takes its sequence with it', () => {
    const { ed, box, button } = animated();
    ed.run(setAnimationTrigger(ed, SLIDE, 0, button));
    const back = openEditor(ed.save());
    expect(effects(back)).toEqual([[box, button]]);
    ed.run(deleteShapes(ed, SLIDE, [ed.shape(SLIDE, button)!]));
    expect(effects(ed)).toEqual([]);
  });

  test('a shape that is not on the slide cannot be a trigger, and nothing changes', () => {
    const { ed, box } = animated();
    expect(() => ed.run(setAnimationTrigger(ed, SLIDE, 0, 9999))).toThrow(/no shape 9999/);
    expect(effects(ed)).toEqual([[box, undefined]]);
  });
});

describe('setShapeAnimationTrigger', () => {
  test('every effect on the shape moves, whatever sequence each was in, as one undo step; other shapes\' effects stay', () => {
    const { ed, box, button } = animated();
    ed.run(addAnimation(ed, SLIDE, ed.shape(SLIDE, button)!, fade));
    ed.run(addAnimation(ed, SLIDE, ed.shape(SLIDE, box)!, findPreset('exit', 10)!));
    ed.run(addAnimation(ed, SLIDE, ed.shape(SLIDE, box)!, findPreset('emph', 8)!, { triggerShape: 2 }));
    ed.log.clear();
    expect(effects(ed)).toEqual([[box, undefined], [button, undefined], [box, undefined], [box, 2]]);
    expect(cycle(ed, () => setShapeAnimationTrigger(ed, SLIDE, box, button))).toEqual([]);
    ed.run(setShapeAnimationTrigger(ed, SLIDE, box, button));
    expect(effects(ed)).toEqual([[button, undefined], [box, button], [box, button], [box, button]]);
    ed.run(setShapeAnimationTrigger(ed, SLIDE, box, null));
    expect(effects(ed)).toEqual([[button, undefined], [box, undefined], [box, undefined], [box, undefined]]);
    ed.undo();
    ed.undo();
    expect(effects(ed)).toEqual([[box, undefined], [button, undefined], [box, undefined], [box, 2]]);
  });
});

