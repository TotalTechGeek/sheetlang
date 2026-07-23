/**
 * Generates `src/parser.js` from `src/grammar.peggy`.
 *
 * The parser is generated ahead of time rather than at import time so the runtime
 * has no filesystem dependency and bundles cleanly.
 */

import { readFile, writeFile } from 'node:fs/promises'
import { fileURLToPath } from 'node:url'
import { dirname, join } from 'node:path'
import peggy from 'peggy'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const grammarPath = join(root, 'src', 'grammar.peggy')
const outputPath = join(root, 'src', 'parser.js')

const grammar = await readFile(grammarPath, 'utf8')

const source = peggy.generate(grammar, {
  output: 'source',
  format: 'es',
  grammarSource: 'src/grammar.peggy',
  allowedStartRules: ['Formula']
})

const banner = `// GENERATED FILE — do not edit.
// Produced from src/grammar.peggy by \`npm run build\`.
/* eslint-disable */

`

await writeFile(outputPath, banner + source)
console.log(`wrote ${outputPath} (${(source.length / 1024).toFixed(1)} kB)`)
