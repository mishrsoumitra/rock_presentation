// tools/start/start.test.ts  -- the pieces of npm start that decide what runs (tools/start/start.ts)
import { existsSync, mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { basename, dirname, join, resolve } from 'node:path';
import { describe, expect, test } from 'vitest';
import { chromiumShell, filesUnder, hashFiles, playwrightBrowsersDir, satisfies } from './start.ts';

describe('npm start', () => {
  test('checks the Node version against package.json engines', () => {
    expect(satisfies('24.14.0', '>=24.11.0')).toBe(true);
    expect(satisfies('24.11.0', '>=24.11.0')).toBe(true);
    expect(satisfies('24.10.9', '>=24.11.0')).toBe(false);
    expect(satisfies('25.0.0', '>=24.11.0')).toBe(true);
    expect(satisfies('9.99.99', '>=24.11.0')).toBe(false);
    expect(() => satisfies('24.14.0', '^24')).toThrow(/unsupported/);
  });

  test("a step's stamp is a hash of its inputs: the same whatever their order, different after any change", () => {
    const dir = mkdtempSync(join(tmpdir(), 'pptx-start-'));
    try {
      const a = join(dir, 'a.txt');
      const b = join(dir, 'sub', 'b.txt');
      mkdirSync(dirname(b));
      writeFileSync(a, 'a');
      writeFileSync(b, 'b');
      writeFileSync(join(dir, '.hidden'), 'not an input');
      const before = hashFiles([a, b], dir);
      expect(hashFiles([b, a], dir)).toBe(before);
      expect(before).toMatch(/^[0-9a-f]{64}$/);
      writeFileSync(b, 'changed');
      expect(hashFiles([a, b], dir)).not.toBe(before);
      // a file that is not there counts as empty, so a missing input still hashes, and differs from no input
      expect(hashFiles([a, join(dir, 'missing')], dir)).not.toBe(hashFiles([a], dir));
      expect(filesUnder(dir).map(p => basename(p)).sort()).toEqual(['a.txt', 'b.txt']);
      expect(filesUnder(join(dir, 'nowhere'))).toEqual([]);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  });

  test("Playwright's browsers directory, by its own rules", () => {
    const core = '/pkg/playwright-core';
    expect(playwrightBrowsersDir(core, { PLAYWRIGHT_BROWSERS_PATH: '0' })).toBe(join(core, '.local-browsers'));
    expect(playwrightBrowsersDir(core, { PLAYWRIGHT_BROWSERS_PATH: '/srv/browsers' })).toBe(resolve('/srv/browsers'));
    expect(playwrightBrowsersDir(core, { PLAYWRIGHT_BROWSERS_PATH: 'browsers', INIT_CWD: '/work' })).toBe(resolve('/work', 'browsers'));
    expect(playwrightBrowsersDir(core, {}, 'darwin', '/Users/me')).toBe(join('/Users/me', 'Library', 'Caches', 'ms-playwright'));
    expect(playwrightBrowsersDir(core, {}, 'linux', '/home/me')).toBe(join('/home/me', '.cache', 'ms-playwright'));
    expect(playwrightBrowsersDir(core, { XDG_CACHE_HOME: '/var/cache/me' }, 'linux', '/home/me')).toBe(join('/var/cache/me', 'ms-playwright'));
    expect(playwrightBrowsersDir(core, { LOCALAPPDATA: 'C:\\Users\\me\\AppData\\Local' }, 'win32', 'C:\\Users\\me')).toBe(join('C:\\Users\\me\\AppData\\Local', 'ms-playwright'));
    expect(() => playwrightBrowsersDir(core, {}, 'freebsd', '/home/me')).toThrow(/freebsd/);
  });

  test('the headless shell the server draws with: a Playwright browser directory, and the CLI that installs it', () => {
    const shell = chromiumShell({});
    expect(basename(shell.dir)).toMatch(/^chromium_headless_shell-\d+$/);
    expect(basename(dirname(shell.dir))).toBe('ms-playwright');
    expect(shell.version).toMatch(/^\d+\./);
    expect(existsSync(shell.cli)).toBe(true);
  });
});
