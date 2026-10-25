// packages/opc/src/names.ts  -- part names and relationship targets (ECMA-376 Part 2, sections 6.2 and 6.5)

/** '/ppt/slides/slide1.xml' from the ZIP item name 'ppt/slides/slide1.xml', or from an already absolute name. */
export const normalizeName = (n: string) => (n.startsWith('/') ? n : '/' + n);

export const asciiLower = (s: string) => s.replace(/[A-Z]/g, c => String.fromCharCode(c.charCodeAt(0) + 32));

/** Part names are equivalent when they match as case-insensitive ASCII strings; percent-escapes are decoded first
 *  so a relationship target 'media/image%201.png' finds the ZIP item 'ppt/media/image 1.png'. */
export function partKey(name: string): string {
  let n = normalizeName(name);
  if (n.includes('%')) {
    try {
      n = decodeURIComponent(n);
    } catch {
      // a malformed escape: compare the name as written
    }
  }
  return asciiLower(n);
}

/** Lower-case extension of a part name, '' when the last segment has none. */
export function extensionOf(partName: string): string {
  const last = partName.slice(partName.lastIndexOf('/') + 1);
  const dot = last.lastIndexOf('.');
  return dot < 0 ? '' : asciiLower(last.slice(dot + 1));
}

/** The relationships part of a part: /ppt/slides/slide1.xml -> /ppt/slides/_rels/slide1.xml.rels, and '/' (the package) -> /_rels/.rels. */
export function relsNameFor(partName: string): string {
  if (partName === '/') return '/_rels/.rels';
  const slash = partName.lastIndexOf('/');
  return `${partName.slice(0, slash)}/_rels/${partName.slice(slash + 1)}.rels`;
}

/** The part a relationships part describes ('/' for the package), or null when the name is not a relationships part name. */
export function sourceOfRels(relsName: string): string | null {
  const m = /^(.*)\/_rels\/([^/]*)\.rels$/i.exec(normalizeName(relsName));
  if (!m) return null;
  const [, folder = '', file = ''] = m;
  if (file === '') return folder === '' ? '/' : null;
  return `${folder}/${file}`;
}

function removeDotSegments(segments: string[]): string {
  const out: string[] = [];
  for (const seg of segments) {
    if (seg === '..') out.pop();                // never above the package root
    else if (seg !== '.' && seg !== '') out.push(seg);
  }
  return '/' + out.join('/');
}

/** Resolve a relationship target against its source part, as RFC 3986 section 5.2 does with the source part name as base. */
export function resolveTarget(sourcePart: string, target: string): string {
  let t = target;
  const hash = t.indexOf('#');
  if (hash >= 0) t = t.slice(0, hash);          // a fragment never names a different part
  if (t.startsWith('/')) return removeDotSegments(t.split('/'));
  const base = sourcePart === '/' ? [] : sourcePart.split('/').slice(1, -1);
  return removeDotSegments([...base, ...t.split('/')]);
}

/** The relative reference PowerPoint writes from one part to another, e.g. '../media/image1.png'. */
export function relativeTarget(sourcePart: string, targetPart: string): string {
  const from = sourcePart === '/' ? [] : sourcePart.split('/').slice(1, -1);
  const to = targetPart.split('/').slice(1);
  let common = 0;
  while (common < from.length && common < to.length - 1 && from[common] === to[common]) common++;
  return [...Array<string>(from.length - common).fill('..'), ...to.slice(common)].join('/');
}

/** Why a ZIP entry name cannot be loaded as a part, or null when it can. Entries ending in '/' are folders and never reach this. */
export function unsafeEntryName(name: string): string | null {
  if (name === '') return 'empty name';
  if (name.startsWith('/')) return 'absolute path';
  if (name.includes('\\')) return 'backslash in name';
  if (/[\0-\x1f\x7f]/.test(name)) return 'control character in name';
  if (name.split('/').some(s => s === '.' || s === '..')) return 'dot segment in name';
  return null;
}
