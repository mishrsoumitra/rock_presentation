// tools/corpus/compare.ts  -- what "lossless" means, shared by the tests, the corpus report and the resave CLI
import { unzipSync } from 'fflate';
import { canonical, canonicalDiff, decodeXmlBytes, parseXml } from '@pptx/xml';
import { DEFAULT_LIMITS } from '@pptx/opc';

export const isXmlEntry = (name: string) => /\.(xml|rels)$/i.test(name);

/** Byte equality without going through expect's element-by-element deep equality, which is slow on megabytes. */
export const sameBytes = (x: Uint8Array, y: Uint8Array) => x.length === y.length && Buffer.compare(x, y) === 0;

function canonicalOf(bytes: Uint8Array): string | null {
  try {
    return canonical(parseXml(decodeXmlBytes(bytes).text));
  } catch {
    return null;   // not well-formed: such a part must come back byte for byte
  }
}

/**
 * fflate's unzip with the limits our own reader applies (entry count, declared sizes), so the tool cannot be made
 * to inflate a bomb. Declared sizes can lie, so this guards tooling; it is not the product's defence (that is
 * packages/opc's readZip, which also counts what it inflates).
 */
function boundedUnzip(bytes: Uint8Array): ReturnType<typeof unzipSync> {
  let entries = 0, total = 0;
  return unzipSync(bytes, {
    filter: f => {
      entries++;
      total += f.originalSize;
      if (entries > DEFAULT_LIMITS.maxEntries || f.originalSize > DEFAULT_LIMITS.maxEntryBytes || total > DEFAULT_LIMITS.maxTotalBytes) {
        throw new Error(`archive exceeds limits at ${f.name}`);
      }
      return true;
    },
  });
}

/**
 * Every way `saved` differs from `original` as PowerPoint would read them, using fflate's unzip as an independent
 * reader: entries missing or added (folder entries aside), XML parts whose canonical form changed, and any other
 * part whose bytes changed. An empty list means the save was lossless.
 */
export function packageDifferences(original: Uint8Array, saved: Uint8Array): string[] {
  const a = boundedUnzip(original);
  const b = boundedUnzip(saved);
  const problems: string[] = [];
  const namesA = Object.keys(a).filter(n => !n.endsWith('/')).sort();
  const namesB = Object.keys(b).sort();
  for (const n of namesA) if (!(n in b)) problems.push(`missing after save: ${n}`);
  for (const n of namesB) if (!(n in a)) problems.push(`added by save: ${n}`);
  for (const n of namesA) {
    const x = a[n]!;
    const y = b[n];
    if (!y) continue;
    const cx = isXmlEntry(n) ? canonicalOf(x) : null;
    if (cx !== null) {
      const cy = canonicalOf(y);
      const diff = cy === null ? 'no longer parses' : canonicalDiff(cx, cy);
      if (diff) problems.push(`${n}: ${diff}`);
    } else if (!sameBytes(x, y)) {
      problems.push(`${n}: bytes differ (${x.length} -> ${y.length})`);
    }
  }
  return problems;
}
