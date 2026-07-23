import { build, evaluate, compile, defaultEngine, Sheet } from '../src/index.js'

function bench (label, iterations, fn) {
  fn(); fn() // warm
  const started = process.hrtime.bigint()
  for (let i = 0; i < iterations; i++) fn(i)
  const ns = Number(process.hrtime.bigint() - started)
  const perOp = ns / iterations
  console.log(
    label.padEnd(46),
    `${(perOp).toFixed(0).padStart(8)} ns/op`,
    `${Math.round(1e9 / perOp).toLocaleString().padStart(14)} ops/sec`
  )
  return perOp
}

console.log('\n=== baseline: plain JS ===')
const rows = Array.from({ length: 1000 }, (_, i) => ({ qty: (i % 9) + 1, price: 19.99, discount: (i % 5) / 100 }))
let k = 0
bench('native JS arithmetic', 1e6, () => {
  const r = rows[k++ % 1000]
  return Math.round(r.qty * r.price * (1 - r.discount) * 100) / 100
})

console.log('\n=== scalar formula ===')
const priced = build('=ROUND(qty * price * (1 - discount), 2)')
k = 0
bench('build() scalar', 1e6, () => priced(rows[k++ % 1000]))
k = 0
bench('engine.run() scalar (interpreted)', 2e5, () => defaultEngine.run(compile('=ROUND(qty*price*(1-discount),2)'), rows[k++ % 1000]))
const scalarLogic = compile('=ROUND(qty * price * (1 - discount), 2)')
k = 0
bench('engine.run() scalar (precompiled logic)', 1e6, () => defaultEngine.run(scalarLogic, rows[k++ % 1000]))
k = 0
bench('evaluate() scalar (re-parses)', 2e4, () => evaluate('=ROUND(qty * price * (1 - discount), 2)', rows[k++ % 1000]))

console.log('\n=== ranges ===')
const grid = {}
for (let r = 1; r <= 10; r++) for (const c of 'ABCDE') grid[c + r] = r
const sum50 = build('=SUM(A1:E10)')
bench('build() SUM over 50-cell range', 2e5, () => sum50(grid))

const bigGrid = {}
for (let r = 1; r <= 200; r++) for (const c of 'ABCDE') bigGrid[c + r] = r
const sum1000 = build('=SUM(A1:E200)')
bench('build() SUM over 1000-cell range', 2e4, () => sum1000(bigGrid))

const lookup = {}
for (let r = 1; r <= 200; r++) { lookup['A' + r] = r; lookup['B' + r] = 'row' + r }
const vlookup = build('=VLOOKUP(150, A1:B200, 2, FALSE)')
bench('build() VLOOKUP over 200 rows', 2e4, () => vlookup(lookup))

console.log('\n=== logical / text ===')
const branch = build('=IF(score >= 90, "A", IF(score >= 80, "B", "C"))')
k = 0
bench('build() nested IF', 1e6, () => branch({ score: k++ % 100 }))
const text = build('=UPPER(LEFT(name, 3)) & "-" & TEXT(n, "000")')
k = 0
bench('build() text pipeline', 3e5, () => text({ name: 'spreadsheet', n: k++ % 1000 }))

console.log('\n=== compile ===')
bench('compile() a short formula', 3e4, () => compile('=SUM(A1:A10) * (1 + taxRate)'))
bench('compile() a long formula', 1e4, () =>
  compile('=IF(AND(a>1,b<2), ROUND(SUM(A1:A10)*rate, 2), IFERROR(VLOOKUP(k, B1:D9, 3, FALSE), "n/a"))'))

console.log('\n=== Sheet ===')
const cells = { A1: 1 }
for (let i = 2; i <= 100; i++) cells['A' + i] = `=A${i - 1} + 1`
bench('Sheet: 100-deep chain, cold read', 2e4, () => new Sheet(cells).get('A100'))
const warm = new Sheet(cells)
warm.get('A100')
bench('Sheet: cached read', 1e6, () => warm.get('A100'))
bench('Sheet: set + recalc chain', 2e4, () => { warm.set('A1', Math.random()); return warm.get('A100') })
console.log()

console.log('\n=== incremental recalculation ===')
console.log('Edit one cell, then re-render; only one cell depends on the edit.\n')
for (const n of [100, 1000, 10000]) {
  const sheetCells = { base: 2, A1: '=base * 2' }
  for (let i = 2; i <= n; i++) sheetCells['A' + i] = '=' + i + ' * 3'
  const names = Object.keys(sheetCells).filter((key) => key !== 'base')
  const timings = []
  for (const incremental of [false, true]) {
    const sheet = new Sheet(sheetCells, { incremental })
    for (const name of names) sheet.get(name)
    const reps = 10
    const started = process.hrtime.bigint()
    for (let r = 0; r < reps; r++) {
      sheet.set('base', r)
      for (const name of names) sheet.get(name)
    }
    timings.push(Number(process.hrtime.bigint() - started) / reps / 1e6)
  }
  console.log(
    `${String(n).padStart(6)} cells   whole-sheet ${timings[0].toFixed(2).padStart(6)} ms` +
    `   incremental ${timings[1].toFixed(3).padStart(6)} ms   ${(timings[0] / timings[1]).toFixed(0)}x`
  )
}
console.log()
