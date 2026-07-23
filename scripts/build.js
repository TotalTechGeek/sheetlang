/**
 * Produces the published build in `dist/`:
 *
 *   dist/esm/index.js   ES module
 *   dist/cjs/index.js   CommonJS
 *   dist/types/*.d.ts   hand-written type declarations (copied from ./types)
 *
 * `json-logic-engine` is left external so it resolves to whatever version the
 * consumer installed, rather than being inlined. `peggy` never appears in the
 * output — it is a build-time tool that generates `src/parser.js`, which the
 * runtime imports as plain generated code.
 *
 * The parser is regenerated first, so `npm run build` is a single command.
 */

import { build } from 'esbuild'
import { cpSync, mkdirSync, rmSync, writeFileSync } from 'node:fs'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import { execFileSync } from 'node:child_process'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const entry = join(root, 'src', 'index.js')
const dist = join(root, 'dist')

// 1. Regenerate the parser from the grammar.
execFileSync(process.execPath, [join(root, 'scripts', 'build-parser.js')], { stdio: 'inherit' })

// 2. Fresh output directory.
rmSync(dist, { recursive: true, force: true })
mkdirSync(dist, { recursive: true })

const shared = {
  entryPoints: [entry],
  bundle: true,
  // Only json-logic-engine is a runtime dependency; keep it external.
  external: ['json-logic-engine'],
  platform: 'neutral',
  target: ['es2021', 'node18'],
  logLevel: 'info'
}

// 3. ESM and CJS bundles.
await build({ ...shared, format: 'esm', outfile: join(dist, 'esm', 'index.js') })
await build({ ...shared, format: 'cjs', outfile: join(dist, 'cjs', 'index.js'), platform: 'node' })

// The root package is `"type": "module"`, so a bare `.js` under dist/cjs would be
// loaded as ESM. These per-directory markers pin each format, avoiding the
// dual-package trap without renaming files to .mjs/.cjs.
writeFileSync(join(dist, 'esm', 'package.json'), JSON.stringify({ type: 'module' }, null, 2) + '\n')
writeFileSync(join(dist, 'cjs', 'package.json'), JSON.stringify({ type: 'commonjs' }, null, 2) + '\n')

// 4. Type declarations. Hand-written (see ./types) because the public API types are
//    known precisely, which is more accurate than inferring `any` from JSDoc.
cpSync(join(root, 'types'), join(dist, 'types'), { recursive: true })

console.log('\nbuilt dist/esm, dist/cjs and dist/types')
