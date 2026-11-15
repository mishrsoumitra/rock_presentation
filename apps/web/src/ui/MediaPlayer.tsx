// apps/web/src/ui/MediaPlayer.tsx  -- playing a video or a sound where its picture is, in the editor and in the slide show
import { useEffect, useState } from 'react';
import type { Session } from '../state/store.ts';

/** Where on the page the media plays: the place of its picture. */
export interface PlayerBox { left: number; top: number; width: number; height: number; }

/**
 * The browser's own player over the media's picture, started at once. The file is played from this tab's memory
 * (an address that names it for this page only, withdrawn when the player goes). A video fills the picture's box;
 * a sound shows its controls under its icon.
 */
export function MediaPlayer({ session, part, kind, box, fixed }: { session: Session; part: string; kind: 'video' | 'audio'; box: PlayerBox; fixed?: boolean }) {
  const [url, setUrl] = useState<string>();
  useEffect(() => {
    const p = session.ed.pkg.getPart(part);
    if (!p) return undefined;
    const made = URL.createObjectURL(new Blob([p.bytes as BlobPart], { type: p.contentType || (kind === 'video' ? 'video/mp4' : 'audio/mpeg') }));
    setUrl(made);
    return () => URL.revokeObjectURL(made);
  }, [session, part, kind]);
  if (!url) return null;
  const stop = (e: React.SyntheticEvent) => e.stopPropagation();
  const position = fixed ? 'fixed' : 'absolute';
  if (kind === 'video') {
    return <video className="media-player" data-testid="media-player" src={url} controls autoPlay playsInline
      style={{ position, left: box.left, top: box.top, width: box.width, height: box.height }} onClick={stop} onMouseDown={stop} onDoubleClick={stop} onContextMenu={stop} />;
  }
  return <audio className="media-player audio" data-testid="media-player" src={url} controls autoPlay
    style={{ position, left: box.left + box.width / 2 - 150, top: box.top + box.height + 6, width: 300 }} onClick={stop} onMouseDown={stop} onContextMenu={stop} />;
}
