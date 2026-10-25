// packages/anim/src/presets.ts  -- PowerPoint's effect presets: names, templates and the default variant of each
//
// Templates come from two places. presets.generated.ts holds one per (class, id, subtype) found in decks PowerPoint
// wrote, extracted by tools/anim/gen-presets.ts; the template most of the corpus agrees on wins. The ones below are
// written by hand, for standard presets the corpus lacks, following the patterns PowerPoint uses in the corpus
// (an entrance sets visibility then animates; an exit animates, then hides one millisecond before its end).
import { GENERATED_PRESETS } from './presets.generated.ts';
import { templateDuration, type EffectTemplate } from './template.ts';
import type { PresetClass } from './timing.ts';

export interface Preset extends EffectTemplate {
  presetClass: PresetClass;
  presetID: number;
  presetSubtype: number;
  /** The effect's name as PowerPoint shows it. */
  name: string;
  /** The corpus deck and slide it was taken from, or 'hand-written'. */
  source: string;
}

// ------------------------------------------------------------------------------------------------------ names

const names = (list: string): Map<number, string> =>
  new Map(list.split(';').map(item => {
    const [id, name] = item.split('=');
    return [Number(id), name!.trim()] as [number, string];
  }));

const ENTRANCE = names(
  '1=Appear;2=Fly In;3=Blinds;4=Box;5=Checkerboard;6=Circle;7=Crawl In;8=Diamond;9=Dissolve In;10=Fade;' +
  '11=Flash Once;12=Peek In;13=Plus;14=Random Bars;15=Spiral In;16=Split;17=Stretch;18=Strips;19=Basic Swivel;' +
  '20=Wedge;21=Wheel;22=Wipe;23=Basic Zoom;24=Random Effects;25=Boomerang;26=Bounce;27=Color Typewriter;' +
  '28=Credits;29=Ease In;30=Float;31=Grow & Turn;34=Light Speed;35=Pinwheel;37=Rise Up;38=Swish;39=Thin Line;' +
  '40=Unfold;41=Whip;42=Float In;43=Center Revolve;45=Swivel;47=Float Down;48=Sling;49=Spinner;50=Compress;' +
  '51=Magnify;52=Curve Up;53=Zoom;54=Glide;55=Expand;56=Flip;58=Fold');
const EXIT = new Map([...ENTRANCE, ...names(
  '1=Disappear;2=Fly Out;7=Crawl Out;9=Dissolve Out;12=Peek Out;15=Spiral Out;17=Collapse;29=Ease Out;' +
  '31=Shrink & Turn;37=Sink Down;42=Float Out;47=Float Up;50=Stretchy;52=Curve Down;55=Contract')]);
const EMPHASIS = names(
  '1=Fill Color;2=Change Font;3=Font Color;4=Change Font Size;5=Change Font Style;6=Grow/Shrink;7=Line Color;' +
  '8=Spin;9=Transparency;10=Bold Flash;14=Blast;15=Bold Reveal;16=Brush Color;18=Underline;19=Object Color;' +
  '20=Color Wave;21=Complementary Color;22=Complementary Color 2;23=Contrasting Color;24=Darken;25=Desaturate;' +
  '26=Pulse;27=Color Pulse;28=Grow With Color;30=Lighten;31=Style Emphasis;32=Teeter;33=Vertical Highlight;' +
  '34=Wave;35=Blink;36=Shimmer');
const PATHS = names('0=Custom Path;35=Left;42=Down;63=Right;64=Up');
const MEDIA = names('1=Play;2=Pause;3=Stop');

/** The effect's name as PowerPoint shows it in the animation pane. */
export function presetName(presetClass: PresetClass | undefined, presetID: number | undefined): string {
  if (presetClass === undefined || presetID === undefined) return 'Custom';
  switch (presetClass) {
    case 'entr': return ENTRANCE.get(presetID) ?? `Entrance ${presetID}`;
    case 'exit': return EXIT.get(presetID) ?? `Exit ${presetID}`;
    case 'emph': return EMPHASIS.get(presetID) ?? `Emphasis ${presetID}`;
    case 'path': return PATHS.get(presetID) ?? 'Motion Path';
    case 'mediacall': return MEDIA.get(presetID) ?? 'Media';
    case 'verb': return 'OLE Action';
  }
}

const DIRECTIONS = names(
  '1=From Top;2=From Right;3=From Top-Right;4=From Bottom;6=From Bottom-Right;8=From Left;9=From Top-Left;' +
  '12=From Bottom-Left;16=In;32=Out;21=Vertical In;26=Horizontal In;37=Vertical Out;42=Horizontal Out');

/** The effect option a subtype stands for ("From Bottom", "Horizontal", "2 Spokes"), when it has a name. */
export function subtypeName(presetClass: PresetClass, presetID: number, subtype: number): string | undefined {
  if (presetID === 21 && (presetClass === 'entr' || presetClass === 'exit')) return `${subtype} Spoke${subtype === 1 ? '' : 's'}`;
  if (subtype === 10) return [3, 5, 14, 17].includes(presetID) ? 'Horizontal' : undefined;
  if (subtype === 5) return [3, 5, 14].includes(presetID) ? 'Vertical' : undefined;
  return DIRECTIONS.get(subtype);
}

// ------------------------------------------------------------------------------------------- hand-written

const TGT = '<p:tgtEl><p:spTgt spid="0"/></p:tgtEl>';
const val = (v: string) => (/^-?\d+(\.\d+)?$/.test(v) ? `<p:fltVal val="${v}"/>` : `<p:strVal val="${v}"/>`);
const ctn = (dur: number, fill: boolean, delay?: number) => `<p:cTn dur="${dur}"${fill ? ' fill="hold"' : ''}` +
  (delay === undefined ? '/>' : `><p:stCondLst><p:cond delay="${delay}"/></p:stCondLst></p:cTn>`);
const visibility = (to: 'visible' | 'hidden', delay = 0) =>
  `<p:set><p:cBhvr>${ctn(1, true, delay)}${TGT}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>` +
  `<p:to><p:strVal val="${to}"/></p:to></p:set>`;
const filter = (name: string, dur: number, transition: 'in' | 'out' = 'in') =>
  `<p:animEffect transition="${transition}" filter="${name}"><p:cBhvr>${ctn(dur, false)}${TGT}</p:cBhvr></p:animEffect>`;
const anim = (attrName: string, from: string, to: string, dur: number, base = true) =>
  `<p:anim calcmode="lin" valueType="num"><p:cBhvr${base ? ' additive="base"' : ''}>${ctn(dur, true)}${TGT}` +
  `<p:attrNameLst><p:attrName>${attrName}</p:attrName></p:attrNameLst></p:cBhvr>` +
  `<p:tavLst><p:tav tm="0"><p:val>${val(from)}</p:val></p:tav><p:tav tm="100000"><p:val>${val(to)}</p:val></p:tav></p:tavLst></p:anim>`;
// ppt_w and ppt_h from nothing to full size, or back
const resize = (grow: boolean) =>
  anim('ppt_w', grow ? '0' : '#ppt_w', grow ? '#ppt_w' : '0', 500, false) +
  anim('ppt_h', grow ? '0' : '#ppt_h', grow ? '#ppt_h' : '0', 500, false);
const cmd = (command: string) => `<p:cmd type="call" cmd="${command}"><p:cBhvr>${ctn(1, true)}${TGT}</p:cBhvr></p:cmd>`;
const motion = (path: string) =>
  `<p:animMotion origin="layout" path="${path}" pathEditMode="relative" ptsTypes="AA"><p:cBhvr>${ctn(2000, true)}${TGT}` +
  '<p:attrNameLst><p:attrName>ppt_x</p:attrName><p:attrName>ppt_y</p:attrName></p:attrNameLst></p:cBhvr></p:animMotion>';

/** Where Fly In starts (and Fly Out ends) for each direction subtype. */
const OFFSCREEN: [number, string, string][] = [
  [1, '#ppt_x', '0-#ppt_h/2'], [2, '1+#ppt_w/2', '#ppt_y'], [3, '1+#ppt_w/2', '0-#ppt_h/2'], [4, '#ppt_x', '1+#ppt_h/2'],
  [6, '1+#ppt_w/2', '1+#ppt_h/2'], [8, '0-#ppt_w/2', '#ppt_y'], [9, '0-#ppt_w/2', '0-#ppt_h/2'], [12, '0-#ppt_w/2', '1+#ppt_h/2'],
];

/** Filter effects: presetID, subtype, filter, duration. */
const FILTERS: [number, number, string, number][] = [
  [3, 10, 'blinds(horizontal)', 500], [3, 5, 'blinds(vertical)', 500], [4, 16, 'box(in)', 500],
  [4, 32, 'box(out)', 500], [5, 10, 'checkerboard(across)', 500], [5, 5, 'checkerboard(down)', 500],
  [6, 16, 'circle(in)', 2000], [6, 32, 'circle(out)', 2000], [8, 16, 'diamond(in)', 2000],
  [8, 32, 'diamond(out)', 2000], [9, 0, 'dissolve', 500], [10, 0, 'fade', 500], [12, 1, 'slide(fromTop)', 500],
  [12, 2, 'slide(fromRight)', 500], [12, 4, 'slide(fromBottom)', 500], [12, 8, 'slide(fromLeft)', 500],
  [13, 16, 'plus(in)', 2000], [13, 32, 'plus(out)', 2000], [14, 10, 'randombar(horizontal)', 500],
  [14, 5, 'randombar(vertical)', 500], [16, 21, 'barn(inVertical)', 500], [16, 26, 'barn(inHorizontal)', 500],
  [16, 37, 'barn(outVertical)', 500], [16, 42, 'barn(outHorizontal)', 500], [20, 0, 'wedge', 2000],
  [21, 1, 'wheel(1)', 2000], [21, 2, 'wheel(2)', 2000], [21, 3, 'wheel(3)', 2000], [21, 4, 'wheel(4)', 2000],
  [21, 8, 'wheel(8)', 2000], [22, 1, 'wipe(up)', 500], [22, 2, 'wipe(right)', 500], [22, 4, 'wipe(down)', 500],
  [22, 8, 'wipe(left)', 500],
];

type Draft = Omit<Preset, 'name' | 'duration'>;

function handWritten(): Draft[] {
  const out: Draft[] = [];
  const add = (
    presetClass: PresetClass, presetID: number, presetSubtype: number, behaviors: string, attrs: Record<string, string> = { fill: 'hold' },
  ) => out.push({ presetClass, presetID, presetSubtype, behaviors, attrs, source: 'hand-written' });
  for (const [id, subtype, name, dur] of FILTERS) {
    add('entr', id, subtype, visibility('visible') + filter(name, dur));
    add('exit', id, subtype, filter(name, dur, 'out') + visibility('hidden', dur - 1));
  }
  add('entr', 1, 0, visibility('visible'));
  add('exit', 1, 0, visibility('hidden'));
  for (const [subtype, x, y] of OFFSCREEN) {
    add('entr', 2, subtype, visibility('visible') + anim('ppt_x', x, '#ppt_x', 500) + anim('ppt_y', y, '#ppt_y', 500));
    // PowerPoint writes an exit's variables without the #
    const bare = (s: string) => s.replace(/#/g, '');
    add('exit', 2, subtype, anim('ppt_x', 'ppt_x', bare(x), 500) + anim('ppt_y', 'ppt_y', bare(y), 500) + visibility('hidden', 499));
  }
  add('entr', 53, 16, visibility('visible') + resize(true) + filter('fade', 500));
  add('exit', 53, 16, resize(false) + filter('fade', 500, 'out') + visibility('hidden', 499));
  add('entr', 31, 0, visibility('visible') + resize(true) + anim('style.rotation', '90', '0', 500, false) + filter('fade', 500));
  add('exit', 31, 0, resize(false) + anim('style.rotation', '0', '90', 500, false) + filter('fade', 500, 'out') + visibility('hidden', 499));
  add('emph', 8, 0,
    `<p:animRot by="21600000"><p:cBhvr>${ctn(2000, true)}${TGT}<p:attrNameLst><p:attrName>r</p:attrName></p:attrNameLst></p:cBhvr></p:animRot>`);
  add('emph', 6, 0, `<p:animScale><p:cBhvr>${ctn(2000, true)}${TGT}</p:cBhvr><p:by x="150000" y="150000"/></p:animScale>`);
  add('emph', 26, 0,
    '<p:animEffect transition="out" filter="fade"><p:cBhvr><p:cTn dur="500" tmFilter="0, 0; .2, .5; .8, .5; 1, 0"/>' +
    `${TGT}</p:cBhvr></p:animEffect>` +
    `<p:animScale><p:cBhvr><p:cTn dur="250" autoRev="1" fill="hold"/>${TGT}</p:cBhvr><p:by x="105000" y="105000"/></p:animScale>`);
  add('emph', 9, 0,
    `<p:set><p:cBhvr><p:cTn dur="indefinite"/>${TGT}<p:attrNameLst><p:attrName>style.opacity</p:attrName></p:attrNameLst></p:cBhvr>` +
    '<p:to><p:strVal val="0.5"/></p:to></p:set>' +
    `<p:animEffect filter="image" prLst="opacity: 0.5"><p:cBhvr rctx="IE"><p:cTn dur="indefinite"/>${TGT}</p:cBhvr></p:animEffect>`);
  add('path', 35, 0, motion('M 0 0 L -0.25 0 E'));
  add('path', 42, 0, motion('M 0 0 L 0 0.25 E'));
  add('path', 63, 0, motion('M 0 0 L 0.25 0 E'));
  add('path', 64, 0, motion('M 0 0 L 0 -0.25 E'));
  add('mediacall', 1, 0, cmd('playFrom(0.0)'));
  add('mediacall', 2, 0, cmd('togglePause'));
  add('mediacall', 3, 0, cmd('stop'));
  return out;
}

// ------------------------------------------------------------------------------------------------ catalogue

const key = (c: string, id: number, sub: number) => `${c}|${id}|${sub}`;

function catalogue(): Preset[] {
  const out = new Map<string, Preset>();
  for (const g of GENERATED_PRESETS) {
    const { occurrences: _o, agreeing: _a, ...template } = g;
    out.set(key(g.presetClass, g.presetID, g.presetSubtype), { ...template, name: presetName(g.presetClass, g.presetID) });
  }
  for (const h of handWritten()) {
    const k = key(h.presetClass, h.presetID, h.presetSubtype);
    if (!out.has(k)) out.set(k, { ...h, duration: templateDuration(h.behaviors), name: presetName(h.presetClass, h.presetID) });
  }
  return [...out.values()].sort((a, b) =>
    a.presetClass.localeCompare(b.presetClass) || a.presetID - b.presetID || a.presetSubtype - b.presetSubtype);
}

let cached: Preset[] | undefined;

/** Every preset template known: from the corpus, and by hand where the corpus has none. */
export function presets(): readonly Preset[] {
  cached ??= catalogue();
  return cached;
}

/** The subtype an effect gets when none is chosen: PowerPoint's default option. */
const DEFAULT_SUBTYPE = new Map<string, number>([
  ['entr|2', 4], ['exit|2', 4], ['entr|3', 10], ['exit|3', 10], ['entr|4', 16], ['exit|4', 16], ['entr|5', 10],
  ['exit|5', 10], ['entr|6', 32], ['exit|6', 16], ['entr|8', 32], ['exit|8', 16], ['entr|12', 4], ['exit|12', 4],
  ['entr|13', 32], ['exit|13', 16], ['entr|14', 10], ['exit|14', 10], ['entr|16', 37], ['exit|16', 21],
  ['entr|21', 1], ['exit|21', 1], ['entr|22', 4], ['exit|22', 4], ['entr|53', 16], ['exit|53', 16],
]);

/** A preset by class and id, with its default subtype unless one is given. */
export function findPreset(presetClass: PresetClass, presetID: number, presetSubtype?: number): Preset | undefined {
  const all = presets();
  const sub = presetSubtype ?? DEFAULT_SUBTYPE.get(`${presetClass}|${presetID}`);
  return all.find(p => p.presetClass === presetClass && p.presetID === presetID && (sub === undefined || p.presetSubtype === sub))
    ?? (presetSubtype === undefined ? all.find(p => p.presetClass === presetClass && p.presetID === presetID) : undefined);
}
