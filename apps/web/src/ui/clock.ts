// apps/web/src/ui/clock.ts  -- the presenter view's timer text

/** Time as a speaker reads it: minutes and seconds, with hours once there are any ("0:07", "12:30", "1:02:03"). */
export function clock(ms: number): string {
  const total = Math.max(0, Math.floor(ms / 1000));
  const h = Math.floor(total / 3600), m = Math.floor((total % 3600) / 60), s = total % 60;
  const two = (n: number) => String(n).padStart(2, '0');
  return h ? `${h}:${two(m)}:${two(s)}` : `${m}:${two(s)}`;
}
