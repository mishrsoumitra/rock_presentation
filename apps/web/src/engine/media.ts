// apps/web/src/engine/media.ts  -- what a media file is, and the picture that stands for it on the slide

/** The types of video and sound that can be inserted, by file extension. */
const TYPES: Record<string, { kind: 'video' | 'audio'; type: string }> = {
  mp4: { kind: 'video', type: 'video/mp4' }, m4v: { kind: 'video', type: 'video/mp4' }, mov: { kind: 'video', type: 'video/quicktime' },
  webm: { kind: 'video', type: 'video/webm' }, wmv: { kind: 'video', type: 'video/x-ms-wmv' }, avi: { kind: 'video', type: 'video/x-msvideo' },
  mp3: { kind: 'audio', type: 'audio/mpeg' }, m4a: { kind: 'audio', type: 'audio/mp4' }, wav: { kind: 'audio', type: 'audio/wav' },
  ogg: { kind: 'audio', type: 'audio/ogg' }, wma: { kind: 'audio', type: 'audio/x-ms-wma' }, aac: { kind: 'audio', type: 'audio/aac' },
};

/** What a file is, from its name (the extension decides, as the content type is stored by extension in the deck). */
export function mediaTypeOf(name: string): { kind: 'video' | 'audio'; type: string; ext: string } | undefined {
  const ext = /\.([a-z0-9]+)$/i.exec(name)?.[1]?.toLowerCase();
  const t = ext ? TYPES[ext] : undefined;
  return t && ext ? { ...t, ext } : undefined;
}

export interface Poster { bytes: Uint8Array; width: number; height: number; }

const png = async (canvas: OffscreenCanvas): Promise<Uint8Array> => new Uint8Array(await (await canvas.convertToBlob({ type: 'image/png' })).arrayBuffer());

/** The picture for a video the browser cannot show a frame of: a dark panel with a play triangle. */
async function plainPoster(): Promise<Poster> {
  const c = new OffscreenCanvas(640, 360);
  const g = c.getContext('2d')!;
  g.fillStyle = '#1f1f1f';
  g.fillRect(0, 0, 640, 360);
  g.fillStyle = '#e6e6e6';
  g.beginPath();
  g.moveTo(280, 130); g.lineTo(380, 180); g.lineTo(280, 230);
  g.closePath();
  g.fill();
  return { bytes: await png(c), width: 640, height: 360 };
}

/**
 * The poster of a video: a frame from near its start, at the video's own size, as PowerPoint shows one. A video
 * the browser cannot play (a format it does not know) gets a plain panel instead; it is stored and saved all the same.
 */
export async function videoPoster(bytes: Uint8Array, type: string): Promise<Poster> {
  if (typeof document === 'undefined') return plainPoster();
  const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type }));
  const video = document.createElement('video');
  try {
    video.muted = true;
    video.playsInline = true;
    video.preload = 'auto';
    video.src = url;
    await new Promise<void>((done, fail) => {
      const timer = setTimeout(() => fail(new Error('timed out')), 6000);
      video.onloadeddata = () => { clearTimeout(timer); done(); };
      video.onerror = () => { clearTimeout(timer); fail(new Error('cannot play')); };
    });
    // a little way in, where there is more often a picture than at the very first frame
    if (Number.isFinite(video.duration) && video.duration > 0.2) {
      await new Promise<void>(done => { const t = setTimeout(done, 1500); video.onseeked = () => { clearTimeout(t); done(); }; video.currentTime = Math.min(0.1, video.duration / 2); });
    }
    const w = video.videoWidth, h = video.videoHeight;
    if (!w || !h) return await plainPoster();
    const c = new OffscreenCanvas(w, h);
    c.getContext('2d')!.drawImage(video, 0, 0, w, h);
    return { bytes: await png(c), width: w, height: h };
  } catch {
    return plainPoster();
  } finally {
    video.removeAttribute('src');
    video.load();
    URL.revokeObjectURL(url);
  }
}

/** The picture that stands for a sound: a speaker, as in PowerPoint. */
export async function audioIcon(): Promise<Poster> {
  const c = new OffscreenCanvas(128, 128);
  const g = c.getContext('2d')!;
  g.fillStyle = '#5a5a5a';
  g.beginPath();
  g.moveTo(24, 50); g.lineTo(48, 50); g.lineTo(76, 26); g.lineTo(76, 102); g.lineTo(48, 78); g.lineTo(24, 78);
  g.closePath();
  g.fill();
  g.strokeStyle = '#5a5a5a';
  g.lineWidth = 7;
  g.lineCap = 'round';
  for (const r of [16, 30]) { g.beginPath(); g.arc(78, 64, r, -Math.PI / 3.2, Math.PI / 3.2); g.stroke(); }
  return { bytes: await png(c), width: 128, height: 128 };
}
