// Plan, "Dependency hygiene": the parsing path is the attack surface, so the packages on it may depend only on
// reviewed packages, pinned to exact versions, and their sources may import nothing else (no Node built-ins either:
// they must run unchanged in the browser and in workers). Adding to these lists is a security review decision.
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, test } from 'vitest';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '../..');

const ALLOWED: Record<string, string[]> = {
  xml: [],
  opc: ['@pptx/xml', 'fflate'],
  model: ['@pptx/xml', '@pptx/opc'],
  resolve: ['@pptx/xml', '@pptx/opc', '@pptx/model'],
  geometry: ['@pptx/xml'],
  text: ['harfbuzzjs', 'linebreak', 'bidi-js'],
  metafile: [],
  render: ['@pptx/xml', '@pptx/opc', '@pptx/model', '@pptx/resolve', '@pptx/geometry', '@pptx/text', '@pptx/metafile'],
  charts: ['@pptx/xml', '@pptx/opc', '@pptx/model', '@pptx/resolve', '@pptx/render', '@pptx/text'],
  anim: ['@pptx/xml', '@pptx/opc', '@pptx/model', '@pptx/resolve', '@pptx/render'],
  crypto: ['@pptx/xml', 'cfb'],
  'editor-core': ['@pptx/xml', '@pptx/opc', '@pptx/model', '@pptx/resolve', '@pptx/geometry', '@pptx/anim'],
};

const sourcesOf = (dir: string): string[] =>
  readdirSync(dir).flatMap(n => {
    const p = join(dir, n);
    return statSync(p).isDirectory() ? sourcesOf(p) : p.endsWith('.ts') ? [p] : [];
  });

// static imports and re-exports, side-effect imports, and dynamic imports
const importsOf = (file: string): string[] =>
  [...readFileSync(file, 'utf8').matchAll(/(?:^|\n)\s*(?:import|export)\s[^'"]*?from\s+['"]([^'"]+)['"]|(?:^|\n)\s*import\s+['"]([^'"]+)['"]|import\(\s*['"]([^'"]+)['"]\s*\)/g)]
    .map(m => m[1] ?? m[2] ?? m[3] ?? '');

const packageOf = (specifier: string) => (specifier.startsWith('@') ? specifier.split('/').slice(0, 2).join('/') : specifier.split('/')[0]!);

describe.each(Object.entries(ALLOWED))('packages/%s', (pkg, allowed) => {
  const manifest = JSON.parse(readFileSync(join(ROOT, 'packages', pkg, 'package.json'), 'utf8'));

  test('declares only allowed dependencies, pinned exactly', () => {
    const deps: Record<string, string> = { ...manifest.dependencies, ...manifest.peerDependencies, ...manifest.optionalDependencies };
    for (const [name, version] of Object.entries(deps)) {
      expect(allowed, `${name} is not on the allowlist for packages/${pkg}`).toContain(name);
      if (!version.startsWith('workspace:')) expect(version, `${name} must be pinned to an exact version`).toMatch(/^\d+\.\d+\.\d+$/);
    }
  });

  test('sources import only relative modules and allowed packages', () => {
    for (const file of sourcesOf(join(ROOT, 'packages', pkg, 'src'))) {
      for (const spec of importsOf(file)) {
        if (spec.startsWith('.')) continue;
        expect(spec.startsWith('node:'), `${file} imports ${spec}`).toBe(false);
        expect(allowed, `${file} imports ${spec}`).toContain(packageOf(spec));
      }
    }
  });
});

test('every package under packages/ is covered by the allowlist', () => {
  const packages = readdirSync(join(ROOT, 'packages')).filter(n => statSync(join(ROOT, 'packages', n)).isDirectory());
  expect(packages.filter(p => !(p in ALLOWED))).toEqual([]);
});
