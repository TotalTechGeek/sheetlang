/**
 * formulas vs HyperFormula.
 *
 * Kept in its own package: HyperFormula is GPL-3.0-only, and that licence must not
 * reach the parent project through a shared node_modules.
 *
 * Each benchmark is written to be as fair as the two designs allow, and where a
 * comparison is structurally unfair the output says so rather than hiding it.
 */

import { HyperFormula } from 'hyperformula'
import { build, evaluate, compile, Sheet, functionNames } from '../src/index.js'

const HF_CONFIG = { licenseKey: 'gpl-v3' }

function bench (label, iterations, fn) {
  fn(); fn()
  const started = process.hrtime.bigint()
  for (let i = 0; i < iterations; i++) fn(i)
  return { label, perOp: Number(process.hrtime.bigint() - started) / iterations }
}

function report (title, note, ours, theirs) {
  console.log(`\n\x1b[1m${title}\x1b[0m`)
  if (note) console.log(`\x1b[2m${note}\x1b[0m`)
  const format = (ns) => (ns > 1e6 ? `${(ns / 1e6).toFixed(2)} ms` : ns > 1000 ? `${(ns / 1000).toFixed(2)} µs` : `${ns.toFixed(0)} ns`)
  const width = Math.max(ours.label.length, theirs.label.length)
  console.log(`  ${ours.label.padEnd(width)}  ${format(ours.perOp).padStart(10)}`)
  console.log(`  ${theirs.label.padEnd(width)}  ${format(theirs.perOp).padStart(10)}`)
  const ratio = theirs.perOp / ours.perOp
  const winner = ratio > 1 ? 'formulas' : 'HyperFormula'
  console.log(`  \x1b[1m-> ${winner} faster by ${(ratio > 1 ? ratio : 1 / ratio).toFixed(1)}x\x1b[0m`)
}

console.log('HyperFormula', HyperFormula.version, '| functions:',
  HyperFormula.getRegisteredFunctionNames('enGB').length, 'vs ours:', functionNames.length)

// ---------------------------------------------------------------------------
// 1. Scalar formula over many rows of data — the rules-engine workload
// ---------------------------------------------------------------------------
{
  const rows = Array.from({ length: 1000 }, (_, i) => ({
    qty: (i % 9) + 1, price: 19.99, discount: (i % 5) / 100
  }))

  const priced = build('=ROUND(qty * price * (1 - discount), 2)')
  let i = 0
  const ours = bench('formulas: build() once, run per row', 2e5, () => priced(rows[i++ % 1000]))

  // HyperFormula has no variables, so the row must be written into cells and the
  // result read back. This is the honest cost of using it for this job.
  const hf = HyperFormula.buildFromArray(
    [[1, 19.99, 0, '=ROUND(A1*B1*(1-C1), 2)']], HF_CONFIG
  )
  i = 0
  const theirs = bench('HyperFormula: write cells, read result', 2e4, () => {
    const row = rows[i++ % 1000]
    hf.setCellContents({ sheet: 0, row: 0, col: 0 }, [[row.qty, row.price, row.discount]])
    return hf.getCellValue({ sheet: 0, row: 0, col: 3 })
  })

  report('1. Scalar formula per row of data', 'The rules/pricing workload. HyperFormula is not designed for this.', ours, theirs)
}

// ---------------------------------------------------------------------------
// 2. One-off evaluation of a formula string
// ---------------------------------------------------------------------------
{
  const ours = bench('formulas: evaluate()', 2e4, () =>
    evaluate('=ROUND(qty * price * (1 - discount), 2)', { qty: 3, price: 19.99, discount: 0.1 }))

  const hf = HyperFormula.buildFromArray([[3, 19.99, 0.1]], HF_CONFIG)
  const theirs = bench('HyperFormula: calculateFormula()', 2e4, () =>
    hf.calculateFormula('=ROUND(A1*B1*(1-C1), 2)', 0))

  report('2. One-off formula evaluation (parse + run, no reuse)', 'Both re-parse every call.', ours, theirs)
}

// ---------------------------------------------------------------------------
// 3. Parse/compile only
// ---------------------------------------------------------------------------
{
  const ours = bench('formulas: compile()', 2e4, () =>
    compile('=IF(AND(A1>1,B1<2), ROUND(SUM(A1:A10)*C1, 2), IFERROR(VLOOKUP(D1, B1:D9, 3, FALSE), "n/a"))'))

  const hf = HyperFormula.buildFromArray([[1, 2, 3, 4]], HF_CONFIG)
  const theirs = bench('HyperFormula: calculateFormula()', 1e4, () =>
    hf.calculateFormula('=IF(AND(A1>1,B1<2), ROUND(SUM(A1:A10)*C1, 2), IFERROR(VLOOKUP(D1, B1:D9, 3, FALSE), "n/a"))', 0))

  report('3. Parsing a long formula', 'Not equivalent: ours parses only, theirs parses AND evaluates.', ours, theirs)
}

// ---------------------------------------------------------------------------
// 4. Range aggregation
// ---------------------------------------------------------------------------
{
  const cells = {}
  const grid = []
  for (let row = 1; row <= 200; row++) {
    const line = []
    for (let col = 0; col < 5; col++) {
      cells[String.fromCharCode(65 + col) + row] = row
      line.push(row)
    }
    grid.push(line)
  }

  const sum = build('=SUM(A1:E200)')
  const ours = bench('formulas: SUM(A1:E200)', 2e4, () => sum(cells))

  const hfGrid = grid.map((line) => [...line])
  hfGrid[0][5] = '=SUM(A1:E200)'
  const hf = HyperFormula.buildFromArray(hfGrid, HF_CONFIG)
  // Force a real recalculation each time rather than reading a cached value.
  let n = 0
  const theirs = bench('HyperFormula: SUM(A1:E200)', 2e4, () => {
    hf.setCellContents({ sheet: 0, row: 199, col: 4 }, [[++n]])
    return hf.getCellValue({ sheet: 0, row: 0, col: 5 })
  })

  report('4. SUM over a 1,000-cell range (recalculated each time)', null, ours, theirs)
}

// ---------------------------------------------------------------------------
// 5. Incremental recalculation — edit one cell, re-render the sheet
// ---------------------------------------------------------------------------
for (const size of [1000, 10000]) {
  const cells = { A1: 2 }
  const grid = [[2, '=A1*2']]
  for (let i = 2; i <= size; i++) {
    cells['B' + i] = '=' + i + ' * 3'
    grid.push([null, '=' + i + ' * 3'])
  }
  cells.B1 = '=A1*2'

  const sheet = new Sheet(cells)
  const names = Object.keys(cells)
  for (const name of names) sheet.get(name)
  let r = 0
  const ours = bench('formulas: Sheet', 200, () => {
    sheet.set('A1', ++r)
    let total = 0
    for (const name of names) { const v = sheet.get(name); if (typeof v === 'number') total += v }
    return total
  })

  const hf = HyperFormula.buildFromArray(grid, HF_CONFIG)
  r = 0
  const theirs = bench('HyperFormula', 200, () => {
    hf.setCellContents({ sheet: 0, row: 0, col: 0 }, [[++r]])
    const values = hf.getSheetValues(0)
    let total = 0
    for (const row of values) for (const v of row) if (typeof v === 'number') total += v
    return total
  })

  report(`5. Edit 1 cell + read all values (${size.toLocaleString()} formula cells)`,
    'Only one cell actually depends on the edit.', ours, theirs)
}

// ---------------------------------------------------------------------------
// 6. Sheet construction
// ---------------------------------------------------------------------------
{
  const size = 5000
  const cells = { A1: 1 }
  const grid = [[1, '=A1+1']]
  for (let i = 2; i <= size; i++) {
    cells['B' + i] = `=${i} * 3`
    grid.push([null, `=${i} * 3`])
  }
  cells.B1 = '=A1+1'

  const ours = bench('formulas: new Sheet()', 30, () => new Sheet(cells))
  const theirs = bench('HyperFormula: buildFromArray()', 30, () => HyperFormula.buildFromArray(grid, HF_CONFIG))

  report(`6. Building a ${size.toLocaleString()}-formula sheet`, 'Ours compiles only; HyperFormula also evaluates eagerly.', ours, theirs)
}

// ---------------------------------------------------------------------------
// 7. The thing HyperFormula cannot do at all
// ---------------------------------------------------------------------------
console.log('\n\x1b[1m7. Arbitrary data shapes\x1b[0m')
{
  const data = {
    customer: { address: { city: 'Lisbon', country: 'PT' } },
    orders: [{ total: 120 }, { total: 80 }],
    taxRate: 0.23
  }
  const formula = '=customer.address.city & ": " & TEXT((orders[0].total + orders[1].total) * (1 + taxRate), "#,##0.00")'
  console.log('  formulas:    ', JSON.stringify(evaluate(formula, data)))
  console.log('  HyperFormula: no equivalent — the data model is a grid of cells')
  console.log('\n  compiled AST is plain data:')
  console.log('   ', JSON.stringify(compile('=SUM(A1:A3) * taxRate')))
}
console.log()
