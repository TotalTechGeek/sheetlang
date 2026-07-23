import { compile, evaluate, build, createEngine, Sheet, isError } from '../src/index.js'

const rule = (label) => console.log(`\n\x1b[1m${label}\x1b[0m\n${'─'.repeat(label.length)}`)
const show = (value) => (isError(value) ? value.code : JSON.stringify(value))

// ---------------------------------------------------------------------------
rule('Formulas compile straight to JSON Logic')

for (const formula of [
  '=1 + 2 * 3',
  '=SUM(A1:A10) * (1 + taxRate)',
  '=IF(score >= 90, "A", IF(score >= 80, "B", "C"))',
  '=customer.address.city & ", " & customer.address.country',
  '=VLOOKUP(sku, Prices!A1:C50, 3, FALSE)'
]) {
  console.log(formula.padEnd(52), JSON.stringify(compile(formula)))
}

// ---------------------------------------------------------------------------
rule('...and that logic is just data you can run')

const data = {
  A1: 100, A2: 250, A3: 75,
  taxRate: 0.2,
  score: 85,
  customer: { address: { city: 'Lisbon', country: 'PT' } }
}

for (const formula of [
  '=SUM(A1:A3) * (1 + taxRate)',
  '=IF(score >= 90, "A", IF(score >= 80, "B", "C"))',
  '=customer.address.city & ", " & customer.address.country',
  '=TEXT(SUM(A1:A3), "#,##0.00")',
  '=AVERAGE(A1:A3) > 100'
]) {
  console.log(formula.padEnd(52), '=>', show(evaluate(formula, data)))
}

// ---------------------------------------------------------------------------
rule('Errors are values, exactly as in a spreadsheet')

for (const formula of ['=1/0', '=IFERROR(1/0, "fallback")', '="abc" + 1', '=SQRT(-1)', '=IF(B1 = 0, 0, A1/B1)']) {
  console.log(formula.padEnd(52), '=>', show(evaluate(formula, { A1: 5, B1: 0 })))
}

// ---------------------------------------------------------------------------
rule('A sheet wires cells together')

const invoice = new Sheet({
  VAT: 0.2,
  B2: 3, C2: 19.99, D2: '=B2*C2',
  B3: 1, C3: 249.0, D3: '=B3*C3',
  B4: 5, C4: 4.5, D4: '=B4*C4',
  D6: '=SUM(D2:D4)',
  D7: '=ROUND(D6 * VAT, 2)',
  D8: '=ROUND(D6 + D7, 2)',
  Summary: '="Total due: " & TEXT(D8, "#,##0.00")'
})

console.log('subtotal   ', invoice.get('D6'))
console.log('VAT        ', invoice.get('D7'))
console.log('total      ', invoice.get('D8'))
console.log('summary    ', invoice.get('Summary'))

invoice.set('B3', 2)
console.log('\nafter changing B3 to 2:')
console.log('total      ', invoice.get('D8'))

// ---------------------------------------------------------------------------
rule('Circular references are caught, not hung on')

console.log(show(new Sheet({ A1: '=B1 + 1', B1: '=A1 + 1' }).get('A1')))

// ---------------------------------------------------------------------------
rule('The library is extensible')

const engine = createEngine({
  functions: {
    SLUGIFY: {
      method: ([text]) => String(text).toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, ''),
      deterministic: true
    }
  }
})

const formula = '=SLUGIFY(title) & "-" & id'
console.log(formula.padEnd(52), '=>',
  show(evaluate(formula, { title: 'Hello, Wide World!', id: 12 }, { engine, functions: null })))

// ---------------------------------------------------------------------------
rule('Compiled formulas are fast to re-run')

const priced = build('=ROUND(qty * price * (1 - discount), 2)')
const rows = Array.from({ length: 200000 }, (_, i) => ({
  qty: (i % 9) + 1, price: 19.99, discount: (i % 5) / 100
}))

const started = performance.now()
let total = 0
for (const row of rows) total += priced(row)
const elapsed = performance.now() - started

console.log(`${rows.length.toLocaleString()} rows in ${elapsed.toFixed(0)}ms ` +
  `(${Math.round(rows.length / (elapsed / 1000)).toLocaleString()} evaluations/sec)`)
console.log('total      ', total.toFixed(2))
console.log()
