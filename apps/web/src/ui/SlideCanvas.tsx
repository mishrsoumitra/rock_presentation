// apps/web/src/ui/SlideCanvas.tsx  -- the slide being edited: drawing, selection, moving, resizing, rotating
import Konva from 'konva';
import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import { Layer, Line, Rect, Shape, Stage, Transformer } from 'react-konva';
import { moveBy, setTransforms, type TransformChange } from '@pptx/editor-core';
import { ShapeView } from '@pptx/model';
import { drawPrimitive, drawShape, drawSlide, shapeTextLayout, type RenderShape, type RenderSlide } from '@pptx/render';
import type { XmlElement } from '@pptx/xml';
import { session, useEditor } from '../state/store.ts';
import { caretColour, caretRect, clickPos, hit, onTextLine, samePos, wordAt, type TextView } from './caret.ts';
import { keystrokeEnd } from './perf.ts';
import { boundsOf, drawBase, parentOf, placementOf, restMatrix, xfrmFor, type Placement } from './placement.ts';
import { promptOf } from './prompts.ts';
import { TextInput } from './TextInput.tsx';
import * as act from './actions.ts';

const ids = new WeakMap<XmlElement, string>();
let nextId = 1;
/** A stable Konva id for a shape element (elements keep their identity across undo and redo). */
export const keyOf = (el: XmlElement) => {
  let id = ids.get(el);
  if (!id) ids.set(el, (id = `s${nextId++}`));
  return id;
};

const MARGIN = 32;
/** How far the mouse moves, in screen pixels, before a press on a selected box drags it rather than clicks it. */
const DRAG_PX = 4;
/** The hit band around a shape's outline, in screen pixels (Konva's hitStrokeWidth, centred on the outline). */
const BAND_PX = 8;

/** Shapes the user can pick: the slide's own, or the members of the group entered. */
function pickable(slide: RenderSlide, entered: XmlElement | null): RenderShape[] {
  const own = slide.shapes.filter(s => s.layer === 'slide' && s.source);
  if (!entered) return own;
  const find = (list: RenderShape[]): RenderShape | undefined => {
    for (const s of list) {
      if (s.source.view.el === entered) return s;
      const inner = find(s.children);
      if (inner) return inner;
    }
    return undefined;
  };
  return find(own)?.children ?? own;
}

export type PlacedTextView = TextView & { transform: [number, number, number, number, number, number] };

/**
 * A shape's text laid out as drawn (also when empty), with the transform to the slide. The render tree already
 * holds the layout of text that shows, so only an empty body is laid out here.
 */
export function textViewOf(rs: RenderShape | undefined): PlacedTextView | undefined {
  const s = session();
  if (!rs || !s || !rs.source?.text || !rs.source.view.txBody) return undefined;
  for (let i = rs.prims.length - 1; i >= 0; i--) {
    const prim = rs.prims[i]!;
    if (prim.kind === 'text') return { txBody: rs.source.view.txBody, resolved: rs.source.text, layout: prim.layout, transform: prim.transform };
  }
  const t = shapeTextLayout(rs.source, s.fonts);
  return t ? { txBody: rs.source.view.txBody, resolved: rs.source.text, layout: t.layout, transform: t.transform } : undefined;
}

/** A point of the slide in a text view's own coordinates. */
function toText(v: PlacedTextView, at: { x: number; y: number }): { x: number; y: number } {
  const inv = invertM(v.transform);
  return { x: inv[0] * at.x + inv[2] * at.y + inv[4], y: inv[1] * at.x + inv[3] * at.y + inv[5] };
}

/**
 * Whether a click at a point of the slide is on a shape's text, so it starts editing there rather than selecting the
 * shape: on a line of its text (a little padded) for a filled shape; anywhere inside for a text box without fill
 * (its hit area is its text already) or an empty placeholder (its prompt). A filled shape with no text is selected.
 */
function onText(rs: RenderShape, v: PlacedTextView, at: { x: number; y: number }, pad: number): boolean {
  const hasText = v.layout.lines.some(l => l.runs.some(r => !r.bullet && r.end > r.start));
  if (!hasText) return !!rs.source.placeholder || rs.source.fill.kind === 'none';
  if (rs.source.fill.kind === 'none') return true;
  const p = toText(v, at);
  return onTextLine(v, p.x, p.y, pad);
}

/**
 * Where a shape without fill can be clicked, besides its drawn outline: the lines of its text, a little padded, in its
 * node's coordinates (its unrotated box). As in PowerPoint, clicks on its empty inside reach the shapes underneath,
 * so text boxes that overlap can each be clicked where their text is. Undefined when the whole box takes clicks:
 * filled shapes, pictures, tables, charts and groups, and shapes with no text (which could not be picked otherwise).
 */
function textHitAreas(rs: RenderShape, p: Placement, pad: number): number[][] | undefined {
  if (rs.source.kind !== 'sp' || rs.source.fill.kind !== 'none') return undefined;
  const t = rs.prims.find(q => q.kind === 'text');
  if (!t || t.kind !== 'text') return undefined;
  const inv = invertM(restMatrix(p));
  const m = t.transform;
  const areas: number[][] = [];
  for (const line of t.layout.lines) {
    if (!line.runs.length) continue;
    const x0 = Math.min(...line.runs.map(g => g.x)) - pad, x1 = Math.max(...line.runs.map(g => g.x + g.advance)) + pad;
    const y0 = line.top - pad, y1 = line.top + line.height + pad;
    areas.push([[x0, y0], [x1, y0], [x1, y1], [x0, y1]].flatMap(([u, v]) => {
      const x = m[0] * u! + m[2] * v! + m[4], y = m[1] * u! + m[3] * v! + m[5];
      return [inv[0] * x + inv[2] * y + inv[4], inv[1] * x + inv[3] * y + inv[5]];
    }));
  }
  return areas.length ? areas : undefined;
}

function findRendered(slide: RenderSlide, el: XmlElement): RenderShape | undefined {
  const walk = (list: RenderShape[]): RenderShape | undefined => {
    for (const s of list) {
      if (s.source?.view.el === el && s.layer === 'slide') return s;
      const inner = walk(s.children);
      if (inner) return inner;
    }
    return undefined;
  };
  return walk(slide.shapes);
}

interface Guide { x?: number; y?: number; }

/**
 * What a mouse-down on a shape started, until the mouse comes up. 'text': the mouse selects text in the box being
 * edited as it moves (tracked on the window, so it goes on outside the box). 'click': the shape was selected; the
 * mouse moving more than DRAG_PX before it comes up drags the shape, and a plain click on its text (`edit`) starts
 * editing there when it comes up.
 */
interface Press {
  kind: 'text' | 'click';
  el: XmlElement;
  rs: RenderShape;
  node: Konva.Node;
  down: MouseEvent;
  at: { x: number; y: number };
  edit: boolean;
}

export function SlideCanvas() {
  const sess = useEditor(s => s.session);
  const part = useEditor(s => s.slide);
  useEditor(s => s.version);
  const selection = useEditor(s => s.selection);
  const entered = useEditor(s => s.entered);
  const text = useEditor(s => s.text);
  const zoomSetting = useEditor(s => s.zoom);
  const drawing = useEditor(s => s.drawing);
  const set = useEditor(s => s.setState);
  const host = useRef<HTMLDivElement>(null);
  const stageRef = useRef<Konva.Stage>(null);
  const trRef = useRef<Konva.Transformer>(null);
  const [size, setSize] = useState({ w: 800, h: 600 });
  const [marquee, setMarquee] = useState<{ x0: number; y0: number; x1: number; y1: number } | null>(null);
  const [guides, setGuides] = useState<Guide[]>([]);
  const drag = useRef<{ starts: Map<string, { x: number; y: number }>; key: string } | null>(null);
  const press = useRef<Press | null>(null);
  const endPress = useRef<() => void>(() => undefined);
  /** The mouse button is down in the text: the caret stays lit instead of blinking, as in other editors. */
  const [holding, setHolding] = useState(false);

  useLayoutEffect(() => {
    const el = host.current;
    if (!el) return;
    const ro = new ResizeObserver(() => setSize({ w: el.clientWidth, h: el.clientHeight }));
    ro.observe(el);
    setSize({ w: el.clientWidth, h: el.clientHeight });
    return () => ro.disconnect();
  }, []);

  // a keystroke's cost is measured to here: the slide has been rendered and committed with the text it typed
  useLayoutEffect(() => { keystrokeEnd(); });
  useEffect(() => () => endPress.current(), []);

  const scene = sess && part ? sess.scenes.get(part) : undefined;
  const slide = scene?.slide;
  useEffect(() => { if (slide) sess?.images.request(slide.images); });

  const cx = slide?.size.cx ?? 12192000, cy = slide?.size.cy ?? 6858000;
  const fit = Math.max(0.00001, Math.min((size.w - 2 * MARGIN) / cx, (size.h - 2 * MARGIN) / cy));
  const zoom = zoomSetting === 'fit' ? fit : zoomSetting;
  const stageW = Math.max(size.w, cx * zoom + 2 * MARGIN), stageH = Math.max(size.h, cy * zoom + 2 * MARGIN);
  const ox = (stageW - cx * zoom) / 2, oy = (stageH - cy * zoom) / 2;
  const pixelRatio = typeof window === 'undefined' ? 1 : window.devicePixelRatio || 1;
  const device = zoom * pixelRatio;
  const createCanvas = (w: number, h: number) => new OffscreenCanvas(w, h);

  const shapes = useMemo(() => (slide ? pickable(slide, entered) : []), [slide, entered]);
  const under = useMemo<RenderSlide | undefined>(() => slide && {
    ...slide,
    // the master's and layout's shapes, and slide shapes outside the group being edited, are drawn but not picked
    shapes: slide.shapes.filter(s => s.layer !== 'slide' || (entered !== null && !shapes.includes(s) && !s.children.some(c => shapes.includes(c)))),
  }, [slide, shapes, entered]);
  const enteredGroup = entered && slide ? findRendered(slide, entered) : undefined;

  // for end-to-end tests and debugging: where slide coordinates land on screen
  useEffect(() => { (globalThis as unknown as { __stage?: Konva.Stage | null }).__stage = stageRef.current; });

  // attach the transformer to the selected nodes
  useEffect(() => {
    const stage = stageRef.current, tr = trRef.current;
    if (!stage || !tr) return;
    const nodes = text ? [] : selection.map(el => stage.findOne(`#${keyOf(el)}`)).filter((n): n is Konva.Node => !!n);
    tr.nodes(nodes);
    tr.getLayer()?.batchDraw();
  });

  const rsOf = (el: XmlElement) => shapes.find(s => s.source.view.el === el);

  if (!sess || !part || !slide) return <div ref={host} className="canvas-host empty">{sess ? 'This presentation has no slides.' : null}</div>;
  const ed = sess.ed;

  /** The pointer in slide coordinates (EMU). */
  const pointer = () => stageRef.current?.getRelativePointerPosition() ?? { x: 0, y: 0 };

  const selectShape = (el: XmlElement, additive: boolean) => {
    const st = useEditor.getState();
    if (additive) {
      const has = st.selection.includes(el);
      st.select(has ? st.selection.filter(x => x !== el) : [...st.selection, el]);
    } else if (!st.selection.includes(el)) st.select([el]);
  };

  /** The text position for a click at a point of the slide. */
  const posAt = (v: PlacedTextView, at: { x: number; y: number }) => {
    const p = toText(v, at);
    return clickPos(v, p.x, p.y);
  };

  /** A box's text as it lays out now: its tree may have been rebuilt since the mouse went down on it. */
  const viewNow = (el: XmlElement): PlacedTextView | undefined => {
    const st = useEditor.getState();
    const sc = st.session && st.slide ? st.session.scenes.get(st.slide) : undefined;
    return sc ? textViewOf(findRendered(sc.slide, el)) : undefined;
  };

  const startText = (rs: RenderShape, at?: { x: number; y: number }) => {
    const v = textViewOf(rs);
    if (!v) return;
    const pos = at ? posAt(v, at) : { p: 0, offset: 0 };
    ed.log.seal();
    set({ text: { part, el: rs.source.view.el, caret: pos, anchor: pos }, selection: [rs.source.view.el], pending: null });
  };

  /** Track the mouse on the window until it comes up: outside the box, and outside the stage, as well as in it. */
  const beginPress = (pr: Press) => {
    endPress.current();
    press.current = pr;
    if (pr.kind === 'text') setHolding(true);
    const move = (ev: MouseEvent) => {
      const p = press.current, stage = stageRef.current;
      if (!p || !stage) return;
      if (p.kind === 'text') {
        const t = useEditor.getState().text;
        if (!t || t.el !== p.el) { endPress.current(); return; }
        stage.setPointersPositions(ev);
        const at = stage.getRelativePointerPosition();
        const v = at ? viewNow(p.el) : undefined;
        if (!at || !v) return;
        const caret = posAt(v, at);
        if (!samePos(caret, t.caret)) useEditor.setState({ text: { ...t, caret } });
        return;
      }
      if (Math.abs(ev.clientX - p.down.clientX) <= DRAG_PX && Math.abs(ev.clientY - p.down.clientY) <= DRAG_PX) return;
      // a drag: the shape (and the rest of the selection) follows the mouse from where it went down
      endPress.current();
      stage.setPointersPositions(p.down);
      p.node.startDrag();
    };
    const up = () => {
      const p = press.current;
      endPress.current();
      if (p?.kind === 'click' && p.edit) startText(p.rs, p.at);
    };
    endPress.current = () => {
      press.current = null;
      endPress.current = () => undefined;
      window.removeEventListener('mousemove', move);
      window.removeEventListener('mouseup', up);
      setHolding(false);
    };
    window.addEventListener('mousemove', move);
    window.addEventListener('mouseup', up);
  };

  /**
   * What is under the mouse on a shape: `v`, its text as laid out (a shape with a body); `onBand`, the band around
   * its outline, which grabs the box (except where a line of text runs right up to the edge, in a box without
   * insets: a press on the text is a press on the text); and `text`, whether a click there starts editing rather
   * than selecting the shape (on a line of its text, or anywhere in a text box or an empty placeholder).
   */
  const probe = (rs: RenderShape, p: Placement, node: Konva.Node, at: { x: number; y: number }) => {
    const local = node.getRelativePointerPosition() ?? { x: p.w / 2, y: p.h / 2 };
    const v = rs.source.kind === 'sp' ? textViewOf(rs) : undefined;
    const band = BAND_PX / 2 / zoom;
    const outside = local.x < 0 || local.y < 0 || local.x > p.w || local.y > p.h;
    const nearEdge = local.x < band || local.y < band || local.x > p.w - band || local.y > p.h - band;
    const q = v ? toText(v, at) : undefined;
    const onBand = outside || (nearEdge && !(v && q && onTextLine(v, q.x, q.y, 0)));
    return { v, onBand, text: !!v && !onBand && onText(rs, v, at, band) };
  };

  /**
   * The mouse goes down on a shape. As in Google Slides: one click on the text of a box puts the caret there and
   * starts editing (an empty placeholder, at its start); in the box being edited the caret goes to the click, Shift
   * extends the selection, and the mouse held down selects text as it moves. The border band, a shape without text,
   * or a box already selected but not being edited: the shape is selected, and the mouse moving before it comes up
   * drags it. A plain click on text never moves a box.
   */
  const onShapeDown = (rs: RenderShape, p: Placement, e: Konva.KonvaEventObject<MouseEvent>) => {
    if (useEditor.getState().drawing) return;          // drawing a new shape over this one
    e.cancelBubble = true;
    const el = rs.source.view.el;
    const st = useEditor.getState();
    const evt = e.evt;
    const node = e.target;
    if (evt.button !== 0) {
      selectShape(el, false);
      return;
    }
    const at = pointer();
    const { v, onBand, text } = probe(rs, p, node, at);
    const editingThis = st.text?.el === el;
    if (editingThis && v && !onBand) {
      evt.preventDefault();                               // the text input keeps the focus
      ed.log.seal();
      const t = st.text!;
      const pos = posAt(v, at);
      set({ text: { ...t, caret: pos, anchor: evt.shiftKey ? t.anchor : pos }, pending: null });
      beginPress({ kind: 'text', el, rs, node, down: evt, at, edit: false });
      return;
    }
    if (evt.shiftKey || evt.metaKey || evt.ctrlKey) {
      selectShape(el, true);                              // adds to or takes from the selection (editing ends)
      return;
    }
    const selectedAlone = st.selection.includes(el) && !st.text;
    if (text && !selectedAlone) {
      evt.preventDefault();
      startText(rs, at);
      beginPress({ kind: 'text', el, rs, node, down: evt, at, edit: false });
      return;
    }
    if (editingThis) {
      ed.log.seal();
      set({ text: null, selection: [el] });
    } else selectShape(el, false);
    beginPress({ kind: 'click', el, rs, node, down: evt, at, edit: text });
  };

  /** Snap the dragged node to slide and shape edges and centres, within a few screen pixels. */
  const snap = (node: Konva.Node, others: RenderShape[]): Guide[] => {
    const tol = 6 / zoom;
    const r = node.getClientRect({ relativeTo: node.getLayer()! });
    const xs = [0, cx / 2, cx], ys = [0, cy / 2, cy];
    for (const o of others) {
      const b = boundsOf(o);
      xs.push(b.x, b.x + b.w / 2, b.x + b.w);
      ys.push(b.y, b.y + b.h / 2, b.y + b.h);
    }
    const found: Guide[] = [];
    const best = (vals: number[], mine: number[]) => {
      let d = Infinity, at = 0;
      for (const v of vals) for (const m of mine) if (Math.abs(v - m) < Math.abs(d)) { d = v - m; at = v; }
      return Math.abs(d) <= tol ? { d, at } : undefined;
    };
    const bx = best(xs, [r.x, r.x + r.width / 2, r.x + r.width]);
    const by = best(ys, [r.y, r.y + r.height / 2, r.y + r.height]);
    if (bx) { node.x(node.x() + bx.d); found.push({ x: bx.at }); }
    if (by) { node.y(node.y() + by.d); found.push({ y: by.at }); }
    return found;
  };

  const onDragStart = (rs: RenderShape) => (e: Konva.KonvaEventObject<DragEvent>) => {
    const stage = stageRef.current!;
    const sel = useEditor.getState().selection.includes(rs.source.view.el) ? useEditor.getState().selection : [rs.source.view.el];
    if (!useEditor.getState().selection.includes(rs.source.view.el)) useEditor.getState().select(sel);
    const starts = new Map<string, { x: number; y: number }>();
    for (const el of sel) {
      const n = stage.findOne(`#${keyOf(el)}`);
      if (n) starts.set(keyOf(el), { x: n.x(), y: n.y() });
    }
    drag.current = { starts, key: e.target.id() };
  };

  const onDragMove = (e: Konva.KonvaEventObject<DragEvent>) => {
    const d = drag.current;
    if (!d) return;
    const stage = stageRef.current!;
    const node = e.target;
    const others = shapes.filter(s => !useEditor.getState().selection.includes(s.source.view.el));
    setGuides(e.evt.altKey ? [] : snap(node, others));
    const start = d.starts.get(node.id())!;
    const dx = node.x() - start.x, dy = node.y() - start.y;
    for (const [id, p] of d.starts) {
      if (id === node.id()) continue;
      const n = stage.findOne(`#${id}`);
      n?.position({ x: p.x + dx, y: p.y + dy });
    }
  };

  const onDragEnd = () => (e: Konva.KonvaEventObject<DragEvent>) => {
    const d = drag.current;
    drag.current = null;
    setGuides([]);
    if (!d) return;
    const start = d.starts.get(e.target.id())!;
    const dx = e.target.x() - start.x, dy = e.target.y() - start.y;
    if (!dx && !dy) return;
    const views = useEditor.getState().selection.map(el => new ShapeView(el, ed.doc.owner(part)));
    useEditor.getState().run(moveBy(ed, part, views, dx, dy));
  };

  const onTransformEnd = () => {
    const tr = trRef.current;
    if (!tr) return;
    const changes: TransformChange[] = [];
    for (const node of tr.nodes()) {
      const el = useEditor.getState().selection.find(x => keyOf(x) === node.id());
      const rs = el && rsOf(el);
      if (!el || !rs) continue;
      const p: Placement = { x: node.x(), y: node.y(), w: node.width() * Math.abs(node.scaleX()), h: node.height() * Math.abs(node.scaleY()), rotation: node.rotation() };
      node.scaleX(1);
      node.scaleY(1);
      changes.push({ shape: new ShapeView(el, ed.doc.owner(part)), xfrm: xfrmFor(p, parentOf(rs), rs.source.xfrm) });
    }
    if (changes.length) useEditor.getState().run(setTransforms(ed, part, changes, 'Resize'));
  };

  const onStageDown = (e: Konva.KonvaEventObject<MouseEvent>) => {
    if (useEditor.getState().drawing) {
      // drawing a new shape: the drag gives its box, wherever it starts
      const p = pointer();
      setMarquee({ x0: p.x, y0: p.y, x1: p.x, y1: p.y });
      return;
    }
    if (e.target !== e.target.getStage() && e.target.name() !== 'underlay') return;
    // empty slide, or the grey around it: text editing ends (its selection with it) and nothing is selected
    const st = useEditor.getState();
    if (st.text) st.setState({ text: null });
    const p = pointer();
    setMarquee({ x0: p.x, y0: p.y, x1: p.x, y1: p.y });
    if (!e.evt.shiftKey) st.select([]);
  };
  const onStageMove = (e: Konva.KonvaEventObject<MouseEvent>) => {
    if (!marquee) return;
    const p = pointer();
    if (useEditor.getState().drawing && e.evt.shiftKey) {
      // Shift keeps it square (or a circle), as in PowerPoint
      const d = Math.max(Math.abs(p.x - marquee.x0), Math.abs(p.y - marquee.y0));
      setMarquee({ ...marquee, x1: marquee.x0 + Math.sign(p.x - marquee.x0 || 1) * d, y1: marquee.y0 + Math.sign(p.y - marquee.y0 || 1) * d });
    } else setMarquee({ ...marquee, x1: p.x, y1: p.y });
  };
  const onStageUp = () => {
    if (!marquee) return;
    const prst = useEditor.getState().drawing;
    if (prst) {
      setMarquee(null);
      useEditor.setState({ drawing: null });
      const w = Math.abs(marquee.x1 - marquee.x0), h = Math.abs(marquee.y1 - marquee.y0);
      const line = /Connector|^line$/.test(prst);
      // a click without a drag gives PowerPoint's default size (1 inch) at the click
      if (w < 4 / zoom && h < 4 / zoom) act.insertShapeHere(prst, { x: marquee.x0, y: marquee.y0, w: 914400, h: line ? 0 : 914400 });
      else act.insertShapeHere(prst, { x: Math.min(marquee.x0, marquee.x1), y: Math.min(marquee.y0, marquee.y1), w, h: line ? (marquee.y1 - marquee.y0) : h });
      return;
    }
    const x0 = Math.min(marquee.x0, marquee.x1), x1 = Math.max(marquee.x0, marquee.x1);
    const y0 = Math.min(marquee.y0, marquee.y1), y1 = Math.max(marquee.y0, marquee.y1);
    setMarquee(null);
    if (x1 - x0 < 2 / zoom && y1 - y0 < 2 / zoom) {
      if (entered) set({ entered: null });
      return;
    }
    // PowerPoint selects what the rectangle encloses entirely
    const inside = shapes.filter(s => {
      const b = boundsOf(s);
      return b.x >= x0 && b.y >= y0 && b.x + b.w <= x1 && b.y + b.h <= y1;
    }).map(s => s.source.view.el);
    const st = useEditor.getState();
    st.select([...new Set([...st.selection, ...inside])]);
  };

  const tv = text ? textViewOf(findRendered(slide, text.el)) : undefined;

  return (
    <div
      ref={host} className={`canvas-host${drawing ? ' drawing' : ''}`} data-testid="canvas"
      onWheel={e => {
        if (!e.ctrlKey && !e.metaKey) return;
        e.preventDefault();
        const next = Math.max(fit / 4, Math.min(fit * 8, zoom * (e.deltaY < 0 ? 1.1 : 1 / 1.1)));
        set({ zoom: next });
      }}
    >
      <Stage
        ref={stageRef} width={stageW} height={stageH} scaleX={zoom} scaleY={zoom} x={ox} y={oy}
        onMouseDown={onStageDown} onMouseMove={onStageMove} onMouseUp={onStageUp}
      >
        <Layer>
          <Shape
            name="underlay" x={0} y={0} width={cx} height={cy} fill="#000"
            sceneFunc={ctx => {
              const c = ctx._context as unknown as CanvasRenderingContext2D;
              const base = c.getTransform();
              if (under) drawSlide(c, under, { scale: device, base: [base.a / device, base.b / device, base.c / device, base.d / device, base.e, base.f], image: p => sess.images.get(p), createCanvas });
            }}
            hitFunc={(ctx, shape) => { ctx.beginPath(); ctx.rect(0, 0, cx, cy); ctx.closePath(); ctx.fillStrokeShape(shape); }}
          />
          {enteredGroup && <Rect {...boundsRect(enteredGroup)} stroke="#7a7a7a" strokeWidth={1 / zoom} dash={[4 / zoom, 3 / zoom]} listening={false} />}
          {shapes.map(rs => {
            const el = rs.source.view.el;
            const p = placementOf(rs);
            const editing = text?.el === el;
            const prompt = editing ? undefined : promptOf(ed, part, rs, sess.fonts);
            // a shape without fill takes clicks on its text and its drawn outline only; a selected box, or the box
            // being typed in, anywhere (its frame can be grabbed to move it, as in PowerPoint)
            const textOnly = editing || selection.includes(el) ? undefined : textHitAreas(rs, p, BAND_PX / 2 / zoom);
            const outlined = !!rs.source.line && rs.source.line.paint.kind !== 'none' && rs.source.line.width > 0;
            return (
              <Shape
                key={keyOf(el)} id={keyOf(el)} x={p.x} y={p.y} width={p.w} height={p.h} offsetX={p.w / 2} offsetY={p.h / 2} rotation={p.rotation}
                // the hit band around the outline is in screen pixels: strokes do not scale with the stage here
                fill="#000" stroke="#000" strokeWidth={0} hitStrokeWidth={BAND_PX} strokeScaleEnabled={false}
                // drags start from the mouse handling above (the border band, or a selected box once the mouse moves)
                draggable={false}
                sceneFunc={ctx => {
                  const c = ctx._context as unknown as CanvasRenderingContext2D;
                  const base = drawBase(c.getTransform(), p, device);
                  drawShape(c, rs, { scale: device, base, image: q => sess.images.get(q), createCanvas });
                  if (rs.source.placeholder && (prompt || editing || !rs.source.text?.paragraphs.some(x => x.runs.some(r => r.text)))) {
                    // empty placeholders: a dashed outline and the prompt, as PowerPoint shows them while editing
                    c.save();
                    c.setTransform(base[0] * device, base[1] * device, base[2] * device, base[3] * device, base[4], base[5]);
                    const m = rs.transform;
                    c.transform(m[0], m[1], m[2], m[3], m[4], m[5]);
                    c.setLineDash([4 / zoom, 3 / zoom]);
                    c.lineWidth = 1 / zoom;
                    c.strokeStyle = 'rgba(120,120,120,0.9)';
                    c.strokeRect(0, 0, rs.box.w, rs.box.h);
                    c.restore();
                    if (prompt) drawPrimitive(c, prompt, { scale: device, base, image: () => undefined });
                  }
                }}
                hitFunc={(ctx, shape) => {
                  ctx.beginPath();
                  if (!textOnly) {
                    ctx.rect(0, 0, p.w, p.h);
                    ctx.closePath();
                    ctx.fillStrokeShape(shape);
                    return;
                  }
                  for (const q of textOnly) {
                    ctx.moveTo(q[0]!, q[1]!);
                    for (let i = 2; i < q.length; i += 2) ctx.lineTo(q[i]!, q[i + 1]!);
                    ctx.closePath();
                  }
                  ctx.fillShape(shape);
                  if (outlined) {
                    ctx.beginPath();
                    ctx.rect(0, 0, p.w, p.h);
                    ctx.strokeShape(shape);
                  }
                }}
                onMouseDown={e => onShapeDown(rs, p, e)}
                onDblClick={e => {
                  e.cancelBubble = true;
                  // Konva counts any two clicks within 400 ms as a double click, wherever they land, and the third
                  // click of a triple click too; the browser's count (the same spot, quickly, the second time) is what
                  // a double click means here: two clicks on different words are two clicks, a third is a triple click
                  if (e.evt.detail !== 2) return;
                  if (rs.source.kind === 'grpSp') { set({ entered: el, selection: [] }); return; }
                  const at = pointer();
                  const v = textViewOf(rs);
                  if (!v) return;
                  if (useEditor.getState().text?.el === el) {
                    // the word under the double click
                    const q = toText(v, at);
                    const w = wordAt(v, hit(v, q.x, q.y));
                    set({ text: { part, el, anchor: w.start, caret: w.end }, pending: null });
                  } else startText(rs, at);
                }}
                onDragStart={onDragStart(rs)} onDragMove={onDragMove} onDragEnd={onDragEnd()}
              />
            );
          })}
        </Layer>
        <Layer>
          {selection.map(el => {
            const rs = rsOf(el);
            if (!rs) return null;
            // the box being typed in gets PowerPoint's editing frame: dashed, a little heavier than a selection's
            return text?.el === el
              ? <Rect key={`e${keyOf(el)}`} {...boundsRect(rs)} stroke="#2b6cb0" strokeWidth={1.5 / zoom} dash={[5 / zoom, 3 / zoom]} listening={false} name="editing-frame" />
              : <Rect key={`o${keyOf(el)}`} {...boundsRect(rs)} stroke="#2b6cb0" strokeWidth={1 / zoom} dash={[3 / zoom, 2 / zoom]} listening={false} />;
          })}
          {tv && text && <TextOverlay view={tv} caret={text.caret} anchor={text.anchor} zoom={zoom} steady={holding} />}
          <Transformer
            ref={trRef} rotationSnaps={[0, 90, 180, 270]} flipEnabled={false} ignoreStroke anchorSize={8} borderStroke="#2b6cb0" anchorStroke="#2b6cb0"
            onTransformEnd={onTransformEnd}
          />
          {guides.map((g, i) => g.x !== undefined
            ? <Line key={`g${i}`} points={[g.x, -cy, g.x, 2 * cy]} stroke="#e53e3e" strokeWidth={1 / zoom} dash={[4 / zoom, 4 / zoom]} listening={false} />
            : <Line key={`g${i}`} points={[-cx, g.y!, 2 * cx, g.y!]} stroke="#e53e3e" strokeWidth={1 / zoom} dash={[4 / zoom, 4 / zoom]} listening={false} />)}
          {marquee && <Rect x={Math.min(marquee.x0, marquee.x1)} y={Math.min(marquee.y0, marquee.y1)} width={Math.abs(marquee.x1 - marquee.x0)} height={Math.abs(marquee.y1 - marquee.y0)}
            fill="rgba(43,108,176,0.08)" stroke="#2b6cb0" strokeWidth={1 / zoom} listening={false} />}
        </Layer>
      </Stage>
      {tv && text && <TextInput view={tv} stage={stageRef} zoom={zoom} />}
      {scene?.error && <div className="slide-error" role="alert">This slide cannot be shown: {scene.error}</div>}
    </div>
  );
}

function boundsRect(rs: RenderShape) {
  const b = boundsOf(rs);
  return { x: b.x, y: b.y, width: b.w, height: b.h };
}

export function invertM(m: readonly number[]): [number, number, number, number, number, number] {
  const [a, b, c, d, e, f] = m as [number, number, number, number, number, number];
  const det = a * d - b * c || 1;
  return [d / det, -b / det, -c / det, a / det, (c * f - d * e) / det, (b * e - a * f) / det];
}

/**
 * The caret and the selection highlight, drawn in the text box's own coordinates, from the editing state alone
 * (the caret and the anchor): the layer redraws in the frame the state changes. With the mouse button down in the
 * text (`steady`) the caret stays lit.
 */
function TextOverlay({ view, caret, anchor, zoom, steady }: {
  view: TextView & { transform: number[] }; caret: { p: number; offset: number }; anchor: { p: number; offset: number }; zoom: number; steady: boolean;
}) {
  const [blink, setBlink] = useState(true);
  useEffect(() => {
    setBlink(true);
    if (steady) return undefined;
    const t = setInterval(() => setBlink(b => !b), 530);
    return () => clearInterval(t);
  }, [caret.p, caret.offset, steady]);
  const m = view.transform;
  return (
    <Shape
      listening={false}
      sceneFunc={ctx => {
        const c = ctx._context as unknown as CanvasRenderingContext2D;
        c.save();
        c.transform(m[0]!, m[1]!, m[2]!, m[3]!, m[4]!, m[5]!);
        const sel = selectionRects(view, anchor, caret);
        c.fillStyle = 'rgba(66, 133, 244, 0.28)';
        for (const r of sel) c.fillRect(r.x, r.top, r.w, r.height);
        const at = caretRect(view, caret);
        if (at && blink && !sel.length) {
          // two screen pixels wide, in the colour of the text it types
          c.fillStyle = caretColour(view, caret);
          c.fillRect(at.x - 1 / zoom, at.top, 2 / zoom, at.height);
        }
        c.restore();
      }}
    />
  );
}

/**
 * Highlight rectangles for a text selection, one per line it touches: from the caret at its start to the caret at
 * its end, whole lines in between (after the bullet, wherever the line is aligned), and a small mark for an empty
 * paragraph inside it.
 */
function selectionRects(v: TextView, a: { p: number; offset: number }, b: { p: number; offset: number }) {
  const [s, e] = a.p < b.p || (a.p === b.p && a.offset <= b.offset) ? [a, b] : [b, a];
  if (s.p === e.p && s.offset === e.offset) return [];
  const rects: { x: number; top: number; w: number; height: number }[] = [];
  const ls = caretRect(v, s), le = caretRect(v, e);
  if (!ls || !le) return rects;
  for (const line of v.layout.lines) {
    if (line.paragraph < s.p || line.paragraph > e.p) continue;
    const top = line.top, height = line.height;
    const lineStart = line.carets[0]?.x ?? line.left, lineEnd = line.carets[line.carets.length - 1]?.x ?? line.left;
    const startsHere = line.paragraph === s.p && ls.top === top;
    const endsHere = line.paragraph === e.p && le.top === top;
    if (line.paragraph === s.p && !startsHere && ls.top > top) continue;
    if (line.paragraph === e.p && !endsHere && le.top < top) continue;
    const x0 = startsHere ? ls.x : lineStart;
    let x1 = endsHere ? le.x : lineEnd;
    if (x1 - x0 < 1 && line.paragraph < e.p) x1 = x0 + height / 4;     // an empty paragraph, or the end of one, inside the selection
    if (x1 > x0) rects.push({ x: x0, top, w: x1 - x0, height });
  }
  return rects;
}
