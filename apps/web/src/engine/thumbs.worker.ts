// apps/web/src/engine/thumbs.worker.ts  -- slide thumbnails, rendered in a worker from its own copy of the deck
//
// The renderer's modules finish loading with a top-level await (the text shaper's WebAssembly). A message that
// reaches a module worker during that await finds no handler and is lost, so this entry installs its handler
// first, keeps what arrives, and loads the renderer afterwards.
import type { ThumbMessage } from './thumbs.render.ts';

export type { ThumbMessage };

const early: MessageEvent<ThumbMessage>[] = [];
let handle: ((e: MessageEvent<ThumbMessage>) => void) | undefined;
self.onmessage = (e: MessageEvent<ThumbMessage>) => (handle ? handle(e) : early.push(e));

const impl = await import('./thumbs.render.ts');
handle = impl.handle;
for (const e of early.splice(0)) handle(e);
