// Presets: the generated catalogue (one template per preset found in decks PowerPoint wrote) and the hand-written
// ones; each must instantiate into a valid effect that reads back as itself and plays as its class says.
import { existsSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, test } from 'vitest';
import { checkOrder } from '@pptx/model';
import { addEffect, findPreset, listEffects, parseTiming, presetName, presets, sample, subtypeName } from '../src/index.ts';
import { GENERATED_PRESETS } from '../src/presets.generated.ts';
import { CORPUS_ROOT } from '../../../tools/corpus/list.ts';
import { renderSlide, shape, slideView, slideXml, spXml } from './helpers.ts';

const all = presets();

describe('the catalogue', () => {
  test('has the presets the corpus uses most, by PowerPoint\'s names', () => {
    const names = (cls: 'entr' | 'exit' | 'emph' | 'path' | 'mediacall', id: number) => [findPreset(cls, id)?.name, presetName(cls, id)];
    expect(names('entr', 2)).toEqual(['Fly In', 'Fly In']);
    expect(names('entr', 10)).toEqual(['Fade', 'Fade']);
    expect(names('entr', 3)).toEqual(['Blinds', 'Blinds']);
    expect(names('entr', 4)).toEqual(['Box', 'Box']);
    expect(names('entr', 1)).toEqual(['Appear', 'Appear']);
    expect(names('exit', 4)).toEqual(['Box', 'Box']);
    expect(names('exit', 10)).toEqual(['Fade', 'Fade']);
    expect(names('entr', 17)).toEqual(['Stretch', 'Stretch']);
    expect(names('entr', 22)).toEqual(['Wipe', 'Wipe']);
    expect(names('path', 42)).toEqual(['Down', 'Down']);
    expect(names('emph', 6)).toEqual(['Grow/Shrink', 'Grow/Shrink']);
    expect(names('emph', 8)).toEqual(['Spin', 'Spin']);
    expect(names('mediacall', 1)).toEqual(['Play', 'Play']);
    expect(subtypeName('entr', 2, 4)).toBe('From Bottom');
    expect(subtypeName('entr', 21, 2)).toBe('2 Spokes');
  });

  test('adds standard presets the corpus lacks, marked as hand-written', () => {
    for (const [cls, id, sub, name] of [['entr', 16, 37, 'Split'], ['entr', 21, 1, 'Wheel'], ['entr', 31, 0, 'Grow & Turn'], ['emph', 26, 0, 'Pulse'],
      ['entr', 2, 9, 'Fly In'], ['exit', 2, 4, 'Fly Out'], ['exit', 22, 4, 'Wipe'], ['mediacall', 3, 0, 'Stop']] as const) {
      const p = findPreset(cls, id, sub);
      expect(p?.name).toBe(name);
      expect(p?.source).toBe('hand-written');
    }
    // where the corpus has one, it wins
    expect(findPreset('entr', 53, 16)?.source).toMatch(/\.pptx#/);
    expect(findPreset('entr', 42)?.source).toMatch(/\.pptx#/);
  });

  test.skipIf(!existsSync(join(CORPUS_ROOT, 'external')))('every generated template names the deck it came from', () => {
    expect(GENERATED_PRESETS.length).toBeGreaterThanOrEqual(44);
    for (const g of GENERATED_PRESETS) {
      const [file] = g.source.split('#');
      expect(existsSync(join(CORPUS_ROOT, 'external', file!)), g.source).toBe(true);
      expect(g.agreeing).toBeGreaterThan(0);
      expect(g.agreeing).toBeLessThanOrEqual(g.occurrences);
    }
  });
});

describe.each(all.map(p => [`${p.presetClass} ${p.presetID}/${p.presetSubtype} ${p.name}`, p] as const))('%s', (_label, preset) => {
  const text = shape(2, 1e6, 1e6, 3e6, 2e6, { paragraphs: ['First', 'Second'] });
  const render = renderSlide([text]);

  test('instantiates into a valid effect that reads back as this preset', () => {
    const { view } = slideView(slideXml(spXml(2, ['First', 'Second'])));
    addEffect(view, 2, preset);
    expect(checkOrder(view.root)).toEqual([]);
    const [e] = listEffects(view);
    expect([e?.presetClass, e?.presetID, e?.presetSubtype, e?.spid]).toEqual([preset.presetClass, preset.presetID, preset.presetSubtype, 2]);
    expect(e?.duration).toBe(preset.duration);
    if (preset.duration > 0) {
      const { view: longer } = slideView(slideXml(spXml(2, ['First', 'Second'])));
      addEffect(longer, 2, preset, { duration: preset.duration * 2 });
      expect(listEffects(longer)[0]?.duration).toBeCloseTo(preset.duration * 2, -1);
    }
  });

  test('plays as its class says: entrances start hidden and end as the slide has it, exits end hidden', () => {
    const { view } = slideView(slideXml(spXml(2, ['First', 'Second'])));
    addEffect(view, 2, preset);
    const t = parseTiming(view.timing!);
    const end = sample(t, render, { clicks: 1, ms: 60_000 });
    if (preset.presetClass === 'entr') {
      expect(sample(t, render, { clicks: 0, ms: 0 }).get(2)).toEqual({ hidden: true });
      expect(end.get(2)).toBeUndefined();
    } else if (preset.presetClass === 'exit') {
      expect(sample(t, render, { clicks: 0, ms: 0 }).has(2)).toBe(false);
      expect(end.get(2)).toEqual({ hidden: true });
    } else {
      for (const ms of [0, preset.duration / 2, preset.duration]) expect(() => sample(t, render, { clicks: 1, ms })).not.toThrow();
    }
  });
});
