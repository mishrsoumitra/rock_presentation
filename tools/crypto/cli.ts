// tools/crypto/cli.ts -- a tiny command line over @pptx/crypto, used by differential.sh to cross-check against
// msoffcrypto-tool. Imported by relative path so it needs no workspace symlink (and so no extra pnpm install).
//
//   node tools/crypto/cli.ts encrypt <in.zip>  <out.bin> <password> [spinCount]
//   node tools/crypto/cli.ts decrypt <in.bin>  <out.zip> <password>
//
// Exit codes: 0 ok, 3 a crypto error (printed as "<code>: <message>"), 64 usage error.
import { readFileSync, writeFileSync } from 'node:fs';
import { decryptPackage, encryptPackage, OfficeCryptoError } from '../../packages/crypto/src/index.ts';

async function main(argv: string[]): Promise<number> {
  const [cmd, input, output, password, spin] = argv;
  if (!cmd || !input || !output || password === undefined) {
    process.stderr.write('usage: cli.ts <encrypt|decrypt> <in> <out> <password> [spinCount]\n');
    return 64;
  }
  const bytes = new Uint8Array(readFileSync(input));
  try {
    const out = cmd === 'encrypt'
      ? await encryptPackage(bytes, password, spin ? { spinCount: Number(spin) } : {})
      : cmd === 'decrypt'
        ? await decryptPackage(bytes, password)
        : null;
    if (out === null) {
      process.stderr.write(`unknown command ${JSON.stringify(cmd)}\n`);
      return 64;
    }
    writeFileSync(output, out);
    return 0;
  } catch (e) {
    if (e instanceof OfficeCryptoError) {
      process.stderr.write(`${e.code}: ${e.message}\n`);
      return 3;
    }
    throw e;
  }
}

main(process.argv.slice(2)).then(code => { process.exitCode = code; });
