// packages/editor-core/src/history.ts  -- commands, the command log, and undo and redo
import { relsNameFor, type OpcPackage, type PackageCheckpoint } from '@pptx/opc';
import type { XmlElement } from '@pptx/xml';
import { capture, restore, type Snapshot } from './snapshot.ts';

export interface Command {
  /** Shown in the Edit menu: "Undo Move". */
  label: string;
  /** Part names the command dirties; also what re-renders and what caches drop. */
  parts: string[];
  do(): void;
  undo(): void;
  /** Consecutive commands with the same key merge into one undo step (nudging with the arrow keys, typing). */
  coalesceKey?: string;
  /** Absorb `next`, which has already run; false keeps it a separate step. */
  coalesce?(next: Command): boolean;
}

export interface EditSpec {
  label: string;
  parts: string[];
  /**
   * The elements the edit may change anything inside of. They must still be in their trees afterwards: to delete a
   * shape, the scope is the shape tree, not the shape. The relationship parts of `parts` are added automatically.
   */
  scope: () => XmlElement[];
  /** The edit adds or removes parts, relationships parts or content types. */
  structural?: boolean;
  apply: () => void;
  coalesceKey?: string;
}

const STATE = Symbol('edit state');

interface EditState {
  roots: XmlElement[];
  before: Snapshot;
  after: Snapshot;
  pkgBefore?: PackageCheckpoint;
  pkgAfter?: PackageCheckpoint;
  /** Dirty flags of the parts before the edit, so undo leaves untouched parts byte-for-byte as they were read. */
  wasDirty: Map<string, boolean>;
  structural: boolean;
}

type EditCommand = Command & { [STATE]?: EditState };

function rootsFor(pkg: OpcPackage, spec: EditSpec): XmlElement[] {
  const roots = spec.scope();
  for (const part of spec.parts) {
    const rels = pkg.getPart(relsNameFor(part));
    if (rels) roots.push(pkg.xml(rels.name).root);
  }
  return roots;
}

/**
 * A command from a function that edits trees in place. Its inverse is exact by construction: before the edit it
 * records the elements in scope (attributes and child lists, not copies of the elements) and, for structural edits,
 * the package's part table; undo puts those back, redo puts back what the edit left. This is how "undo restores a
 * tree that compares equal to the one before" holds for every command without each one writing its own inverse.
 */
export function edit(pkg: OpcPackage, spec: EditSpec): Command {
  const cmd: EditCommand = {
    label: spec.label,
    parts: [...new Set(spec.parts)],
    ...(spec.coalesceKey ? { coalesceKey: spec.coalesceKey } : {}),
    do() {
      const st = cmd[STATE];
      if (st) {                                   // redo
        if (st.pkgAfter) pkg.restore(st.pkgAfter);
        restore(st.after);
        for (const p of cmd.parts) if (pkg.getPart(p)) pkg.markDirty(p);
        return;
      }
      const roots = rootsFor(pkg, spec);
      const structural = spec.structural ?? false;
      const pkgBefore = structural ? pkg.checkpoint() : undefined;
      const wasDirty = new Map(cmd.parts.filter(p => pkg.getPart(p)).map(p => [p, pkg.isDirty(p)] as const));
      const before = capture(roots);
      try {
        spec.apply();
      } catch (e) {
        restore(before);
        if (pkgBefore) pkg.restore(pkgBefore);
        throw e;
      }
      for (const p of cmd.parts) if (pkg.getPart(p)) pkg.markDirty(p);
      const after = capture([...roots, ...rootsFor(pkg, spec)]);
      cmd[STATE] = { roots, before, after, wasDirty, structural, ...(pkgBefore ? { pkgBefore, pkgAfter: pkg.checkpoint() } : {}) };
    },
    undo() {
      const st = cmd[STATE];
      if (!st) return;
      restore(st.before);
      if (st.pkgBefore) pkg.restore(st.pkgBefore);
      else for (const [p, dirty] of st.wasDirty) if (!dirty && pkg.getPart(p)) pkg.getPart(p)!.dirty = false;
    },
    coalesce(next) {
      const a = cmd[STATE];
      const b = (next as EditCommand)[STATE];
      if (!a || !b || a.structural !== b.structural || a.roots.length !== b.roots.length) return false;
      if (a.roots.some((r, i) => r !== b.roots[i])) return false;
      a.after = b.after;
      if (b.pkgAfter) a.pkgAfter = b.pkgAfter;
      for (const [p, d] of b.wasDirty) if (!a.wasDirty.has(p)) a.wasDirty.set(p, d);
      cmd.parts = [...new Set([...cmd.parts, ...next.parts])];
      return true;
    },
  };
  return cmd;
}

/** Several commands as one undo step, undone in reverse order. */
export function compound(label: string, commands: Command[]): Command {
  return {
    label,
    parts: [...new Set(commands.flatMap(c => c.parts))],
    do() { for (const c of commands) c.do(); },
    undo() { for (let i = commands.length - 1; i >= 0; i--) commands[i]!.undo(); },
  };
}

export interface CommandLogOptions {
  /** Commands with the same coalesce key merge only when they follow each other within this time. */
  coalesceWindowMs?: number;
  /** Undo depth; the oldest steps are forgotten. */
  limit?: number;
  now?: () => number;
}

/**
 * The undo and redo stacks. Coalescing differs from a naive "same key as the last command" in three ways: a step
 * that was undone or redone is sealed (typing after an undo starts a new step), steps more than a moment apart do
 * not merge, and the UI can seal explicitly (a click elsewhere ends a typing run).
 */
export class CommandLog {
  private readonly done: { cmd: Command; at: number }[] = [];
  private readonly undone: Command[] = [];
  private sealed = true;
  private readonly window: number;
  private readonly limit: number;
  private readonly now: () => number;
  private readonly pkg: OpcPackage;
  private readonly bump: (parts: string[]) => void;

  constructor(pkg: OpcPackage, bump: (parts: string[]) => void, options: CommandLogOptions = {}) {
    this.pkg = pkg;
    this.bump = bump;
    this.window = options.coalesceWindowMs ?? 1500;
    this.limit = options.limit ?? 500;
    this.now = options.now ?? (() => Date.now());
  }

  run(cmd: Command): void {
    cmd.do();
    const last = this.done[this.done.length - 1];
    const now = this.now();
    if (!this.sealed && last && cmd.coalesceKey !== undefined && last.cmd.coalesceKey === cmd.coalesceKey
      && now - last.at <= this.window && last.cmd.coalesce?.(cmd)) {
      last.at = now;
    } else {
      this.done.push({ cmd, at: now });
      if (this.done.length > this.limit) this.done.shift();
    }
    this.sealed = false;
    this.undone.length = 0;
    this.commit(cmd.parts, true);
  }

  undo(): Command | undefined {
    const e = this.done.pop();
    if (!e) return undefined;
    e.cmd.undo();
    this.undone.push(e.cmd);
    this.sealed = true;
    this.commit(e.cmd.parts, false);
    return e.cmd;
  }

  redo(): Command | undefined {
    const c = this.undone.pop();
    if (!c) return undefined;
    c.do();
    this.done.push({ cmd: c, at: -Infinity });
    this.sealed = true;
    this.commit(c.parts, true);
    return c;
  }

  /** End the current coalescing run: the next command is a new undo step whatever its key. */
  seal(): void { this.sealed = true; }

  get canUndo(): boolean { return this.done.length > 0; }
  get canRedo(): boolean { return this.undone.length > 0; }
  get undoLabel(): string | undefined { return this.done[this.done.length - 1]?.cmd.label; }
  get redoLabel(): string | undefined { return this.undone[this.undone.length - 1]?.label; }

  clear(): void {
    this.done.length = 0;
    this.undone.length = 0;
    this.sealed = true;
  }

  private commit(parts: string[], dirty: boolean): void {
    if (dirty) for (const p of parts) if (this.pkg.getPart(p)) this.pkg.markDirty(p);
    this.bump(parts);
  }
}
