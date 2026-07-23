import test from 'node:test'
import assert from 'node:assert/strict'
import { evaluate, build, isError, defaultEngine } from '../src/index.js'

/** Assert a formula produces a given Excel error token. */
function assertError (formula, code, data) {
  const result = evaluate(formula, data)
  assert.ok(isError(result), `${formula} should be an error, got ${JSON.stringify(result)}`)
  assert.equal(result.code, code)
}

const GRID = { A1: 1, B1: 2, A2: 3, B2: 4, A3: 5, B3: 6 }

test('arithmetic', () => {
  assert.equal(evaluate('=1+2*3'), 7)
  assert.equal(evaluate('=(1+2)*3'), 9)
  assert.equal(evaluate('=10/4'), 2.5)
  assert.equal(evaluate('=2^10'), 1024)
  assert.equal(evaluate('=-2^2'), 4)
  assert.equal(evaluate('=2^3^2'), 64)
  assert.equal(evaluate('=50%'), 0.5)
  assert.equal(evaluate('=200*10%'), 20)
})

test('division by zero yields #DIV/0!', () => {
  assertError('=1/0', '#DIV/0!')
  assertError('=MOD(1,0)', '#DIV/0!')
})

test('blank references coerce to zero', () => {
  assert.equal(evaluate('=missing+5'), 5)
  assert.equal(evaluate('=missing&"x"'), 'x')
})

test('text that is not numeric yields #VALUE!', () => {
  assertError('=A1+1', '#VALUE!', { A1: 'abc' })
  // ...but numeric text converts, as Excel does
  assert.equal(evaluate('=A1+1', { A1: '41' }), 42)
})

test('comparison uses Excel ordering', () => {
  assert.equal(evaluate('="a"="A"'), true, 'text compares case-insensitively')
  assert.equal(evaluate('=EXACT("a","A")'), false, 'EXACT is case-sensitive')
  assert.equal(evaluate('=1<"a"'), true, 'numbers sort before text')
  assert.equal(evaluate('="a"<TRUE'), true, 'text sorts before logicals')
  assert.equal(evaluate('=blank=0'), true, 'a blank compares equal to zero')
  assert.equal(evaluate('=blank=""'), true, 'and equal to empty text')
})

test('concatenation stringifies Excel-style', () => {
  assert.equal(evaluate('="a"&1&TRUE'), 'a1TRUE')
  assert.equal(evaluate('=CONCAT("x", 1, 2)'), 'x12')
})

test('references resolve through val', () => {
  assert.equal(evaluate('=A1+B1', GRID), 3)
  assert.equal(evaluate('=customer.address.city', { customer: { address: { city: 'Ely' } } }), 'Ely')
  assert.equal(evaluate('=orders[1].total', { orders: [{ total: 1 }, { total: 2 }] }), 2)
  assert.equal(evaluate('=[Net Total]', { 'Net Total': 99 }), 99)
  assert.equal(evaluate('=Sheet1!A1', { Sheet1: { A1: 7 } }), 7)
})

test('ranges expand to a grid of values', () => {
  assert.equal(evaluate('=SUM(A1:B3)', GRID), 21)
  assert.equal(evaluate('=SUM(A1:A3)', GRID), 9)
  assert.equal(evaluate('=COUNT(A1:B3)', GRID), 6)
  assert.equal(evaluate('=SUM(Sheet1!A1:A2)', { Sheet1: { A1: 10, A2: 20 } }), 30)
  // Missing cells are blank, not errors
  assert.equal(evaluate('=SUM(A1:A5)', GRID), 9)
})

test('range corners may be given in any order', () => {
  assert.equal(evaluate('=SUM(B3:A1)', GRID), 21)
})

test('IF is lazy, so the untaken branch cannot fail', () => {
  assert.equal(evaluate('=IF(B1=0, 0, A1/B1)', { A1: 5, B1: 0 }), 0)
  assert.equal(evaluate('=IF(B1=0, 0, A1/B1)', { A1: 10, B1: 2 }), 5)
  assert.equal(evaluate('=IF(TRUE, "y")'), 'y')
  assert.equal(evaluate('=IF(FALSE, "y")'), false)
})

test('an error in an IF condition propagates', () => {
  assertError('=IF(1/0>1, "a", "b")', '#DIV/0!')
})

test('IFS and SWITCH', () => {
  assert.equal(evaluate('=IFS(x<10,"low", x<100,"mid", TRUE,"high")', { x: 50 }), 'mid')
  assertError('=IFS(FALSE,1)', '#N/A')
  assert.equal(evaluate('=SWITCH(day, 1, "Mon", 2, "Tue", "other")', { day: 2 }), 'Tue')
  assert.equal(evaluate('=SWITCH(day, 1, "Mon", "other")', { day: 9 }), 'other')
})

test('IFERROR and IFNA', () => {
  assert.equal(evaluate('=IFERROR(1/0, "safe")'), 'safe')
  assert.equal(evaluate('=IFERROR(4/2, "safe")'), 2)
  assert.equal(evaluate('=IFNA(NA(), "gone")'), 'gone')
  assertError('=IFNA(1/0, "gone")', '#DIV/0!')
})

test('logical functions', () => {
  assert.equal(evaluate('=AND(TRUE, 1, "TRUE")'), true)
  assert.equal(evaluate('=AND(TRUE, FALSE)'), false)
  assert.equal(evaluate('=OR(FALSE, 0, 1)'), true)
  assert.equal(evaluate('=NOT(0)'), true)
  assert.equal(evaluate('=XOR(TRUE, TRUE, TRUE)'), true)
})

test('error literals evaluate to errors and are catchable', () => {
  assertError('=#REF!', '#REF!')
  assert.equal(evaluate('=ISERROR(#VALUE!)'), true)
  assert.equal(evaluate('=IFERROR(#NUM!, "ok")'), 'ok')
})

test('errors propagate through arguments', () => {
  assertError('=SUM(1, 1/0)', '#DIV/0!')
  assertError('=UPPER(1/0)', '#DIV/0!')
})

test('math functions', () => {
  assert.equal(evaluate('=ROUND(2.5)'), 3, 'Excel rounds half away from zero')
  assert.equal(evaluate('=ROUND(-2.5)'), -3)
  assert.equal(evaluate('=ROUND(1.005, 2)'), 1.01, 'binary representation is corrected')
  assert.equal(evaluate('=ROUNDUP(1.001, 2)'), 1.01)
  assert.equal(evaluate('=ROUNDDOWN(1.009, 2)'), 1)
  assert.equal(evaluate('=MOD(-3, 5)'), 2, 'MOD takes the sign of the divisor')
  assert.equal(evaluate('=INT(-1.5)'), -2)
  assert.equal(evaluate('=TRUNC(-1.5)'), -1)
  assert.equal(evaluate('=ODD(2)'), 3)
  assert.equal(evaluate('=EVEN(1)'), 2)
  assert.equal(evaluate('=CEILING(4.2, 0.5)'), 4.5)
  assert.equal(evaluate('=FLOOR(4.8, 0.5)'), 4.5)
  assert.equal(evaluate('=GCD(24, 36)'), 12)
  assert.equal(evaluate('=LCM(4, 6)'), 12)
  assert.equal(evaluate('=ABS(-3)'), 3)
  assert.equal(evaluate('=SQRT(16)'), 4)
  assertError('=SQRT(-1)', '#NUM!')
  assert.equal(evaluate('=POWER(2, 8)'), 256)
  assert.equal(Math.round(evaluate('=DEGREES(PI())')), 180)
})

test('statistics', () => {
  assert.equal(evaluate('=AVERAGE(A1:B3)', GRID), 3.5)
  assert.equal(evaluate('=MEDIAN(1, 3, 2)'), 2)
  assert.equal(evaluate('=MEDIAN(1, 2, 3, 4)'), 2.5)
  assert.equal(evaluate('=MAX(A1:B3)', GRID), 6)
  assert.equal(evaluate('=MIN(A1:B3)', GRID), 1)
  assert.equal(evaluate('=LARGE(A1:B3, 2)', GRID), 5)
  assert.equal(evaluate('=SMALL(A1:B3, 2)', GRID), 2)
  assert.equal(evaluate('=STDEV(2,4,4,4,5,5,7,9)').toFixed(4), '2.1381')
  assert.equal(evaluate('=STDEVP(2,4,4,4,5,5,7,9)'), 2)
  assertError('=AVERAGE(x)', '#DIV/0!', { x: null })
})

test('text in a range is skipped by numeric aggregation, unlike direct arguments', () => {
  const data = { A1: 1, A2: 'two', A3: 3 }
  assert.equal(evaluate('=SUM(A1:A3)', data), 4)
  assert.equal(evaluate('=COUNT(A1:A3)', data), 2)
  assert.equal(evaluate('=COUNTA(A1:A3)', data), 3)
  assert.equal(evaluate('=SUM(TRUE, 1)'), 2, 'a direct boolean argument still coerces')
})

test('criteria functions understand operators and wildcards', () => {
  const data = { A1: 5, A2: 15, A3: 25, B1: 'apple', B2: 'apricot', B3: 'banana' }
  assert.equal(evaluate('=COUNTIF(A1:A3, ">10")', data), 2)
  assert.equal(evaluate('=COUNTIF(A1:A3, "<>15")', data), 2)
  assert.equal(evaluate('=COUNTIF(B1:B3, "ap*")', data), 2)
  assert.equal(evaluate('=COUNTIF(B1:B3, "banana")', data), 1)
  assert.equal(evaluate('=SUMIF(A1:A3, ">10")', data), 40)
  assert.equal(evaluate('=SUMIF(B1:B3, "ap*", A1:A3)', data), 20)
  assert.equal(evaluate('=AVERAGEIF(A1:A3, ">=15")', data), 20)
})

test('text functions', () => {
  assert.equal(evaluate('=LEFT("spreadsheet", 6)'), 'spread')
  assert.equal(evaluate('=RIGHT("spreadsheet", 5)'), 'sheet')
  assert.equal(evaluate('=MID("spreadsheet", 7, 5)'), 'sheet')
  assert.equal(evaluate('=LEN("abc")'), 3)
  assert.equal(evaluate('=UPPER("abc")'), 'ABC')
  assert.equal(evaluate('=PROPER("hello wide world")'), 'Hello Wide World')
  assert.equal(evaluate('=TRIM("  a   b  ")'), 'a b')
  assert.equal(evaluate('=SUBSTITUTE("a-b-c", "-", "+")'), 'a+b+c')
  assert.equal(evaluate('=SUBSTITUTE("a-b-c", "-", "+", 2)'), 'a-b+c')
  assert.equal(evaluate('=REPLACE("abcdef", 2, 3, "XY")'), 'aXYef')
  assert.equal(evaluate('=FIND("b", "abcb")'), 2)
  assertError('=FIND("B", "abc")', '#VALUE!')
  assert.equal(evaluate('=SEARCH("B", "abc")'), 2, 'SEARCH is case-insensitive')
  assert.equal(evaluate('=REPT("ab", 3)'), 'ababab')
  assert.equal(evaluate('=TEXTJOIN(", ", TRUE, "a", "", "b")'), 'a, b')
  assert.equal(evaluate('=VALUE("3.5")'), 3.5)
  assert.equal(evaluate('=CHAR(65)'), 'A')
  assert.equal(evaluate('=CODE("A")'), 65)
})

test('TEXT formats numbers and dates', () => {
  assert.equal(evaluate('=TEXT(1234.5, "#,##0.00")'), '1,234.50')
  assert.equal(evaluate('=TEXT(0.256, "0.0%")'), '25.6%')
  assert.equal(evaluate('=TEXT(7, "000")'), '007')
  assert.equal(evaluate('=TEXT(-1234.5, "#,##0.00")'), '-1,234.50')
  assert.equal(evaluate('=TEXT(DATE(2026,7,23), "yyyy-mm-dd")'), '2026-07-23')
  assert.equal(evaluate('=TEXT(DATE(2026,7,23), "d mmmm yyyy")'), '23 July 2026')
})

test('lookup functions', () => {
  const table = { A1: 1, B1: 'one', A2: 2, B2: 'two', A3: 3, B3: 'three' }
  assert.equal(evaluate('=VLOOKUP(2, A1:B3, 2, FALSE)', table), 'two')
  assert.equal(evaluate('=VLOOKUP(2.5, A1:B3, 2, TRUE)', table), 'two', 'approximate match takes the largest value not over')
  assertError('=VLOOKUP(9, A1:B3, 2, FALSE)', '#N/A', table)
  assertError('=VLOOKUP(2, A1:B3, 5, FALSE)', '#REF!', table)
  assert.equal(evaluate('=HLOOKUP(2, A1:B2, 2, FALSE)', { A1: 1, B1: 2, A2: 'x', B2: 'y' }), 'y')
  assert.equal(evaluate('=MATCH(3, A1:A3, 0)', table), 3)
  assert.equal(evaluate('=INDEX(A1:B3, 2, 2)', table), 'two')
  assert.equal(evaluate('=INDEX(A1:A3, 2)', table), 2)
  assert.equal(evaluate('=XLOOKUP(3, A1:A3, B1:B3)', table), 'three')
  assert.equal(evaluate('=XLOOKUP(9, A1:A3, B1:B3, "none")', table), 'none')
  assert.equal(evaluate('=CHOOSE(2, "a", "b", "c")'), 'b')
  assert.equal(evaluate('=ROWS(A1:B3)', table), 3)
  assert.equal(evaluate('=COLUMNS(A1:B3)', table), 2)
})

test('INDEX and MATCH compose the way they do in Excel', () => {
  const table = { A1: 'x', A2: 'y', A3: 'z', B1: 10, B2: 20, B3: 30 }
  assert.equal(evaluate('=INDEX(B1:B3, MATCH("y", A1:A3, 0))', table), 20)
})

test('information functions', () => {
  assert.equal(evaluate('=ISBLANK(nothing)'), true)
  assert.equal(evaluate('=ISNUMBER(1)'), true)
  assert.equal(evaluate('=ISTEXT("a")'), true)
  assert.equal(evaluate('=ISLOGICAL(TRUE)'), true)
  assert.equal(evaluate('=ISERROR(1/0)'), true)
  assert.equal(evaluate('=ISNA(NA())'), true)
  assert.equal(evaluate('=ISERR(NA())'), false, '#N/A is excluded from ISERR')
  assert.equal(evaluate('=ISEVEN(4)'), true)
  assert.equal(evaluate('=ISODD(-3)'), true)
})

test('date functions', () => {
  assert.equal(evaluate('=YEAR(DATE(2026,7,23))'), 2026)
  assert.equal(evaluate('=MONTH(DATE(2026,7,23))'), 7)
  assert.equal(evaluate('=DAY(DATE(2026,7,23))'), 23)
  assert.equal(evaluate('=DAYS(DATE(2026,7,23), DATE(2026,7,1))'), 22)
  assert.equal(evaluate('=YEAR(EDATE(DATE(2026,1,31), 1))'), 2026)
  assert.equal(evaluate('=DAY(EDATE(DATE(2026,1,31), 1))'), 28, 'EDATE clamps to the shorter month')
  assert.equal(evaluate('=DAY(EOMONTH(DATE(2026,2,10), 0))'), 28)
  assert.equal(evaluate('=WEEKDAY(DATE(2026,7,23))'), 5, 'Sunday is 1')
  // Dates behave as serial numbers in arithmetic
  assert.equal(evaluate('=DATE(2026,7,24)-DATE(2026,7,23)'), 1)
})

test('financial functions match Excel to the cent', () => {
  assert.equal(evaluate('=ROUND(PMT(0.05/12, 360, 300000), 2)'), -1610.46)
  assert.equal(evaluate('=ROUND(FV(0.05/12, 120, -500), 2)'), 77641.14)
  assert.equal(evaluate('=ROUND(PV(0.05, 10, -1000), 2)'), 7721.73)
  assert.equal(evaluate('=ROUND(NPV(0.1, 100, 200, 300), 2)'), 481.59)
})

test('build compiles once and reuses the plan', () => {
  const total = build('=SUM(A1:A3) * (1 + tax)')
  assert.equal(total({ A1: 1, A2: 2, A3: 3, tax: 0.1 }).toFixed(2), '6.60')
  assert.equal(total({ A1: 10, A2: 0, A3: 0, tax: 0 }), 10)
})

test('a custom engine can add functions', async () => {
  const { createEngine } = await import('../src/index.js')
  const engine = createEngine({
    functions: { DOUBLE: { method: ([n]) => n * 2, deterministic: true } }
  })
  assert.equal(evaluate('=DOUBLE(21)', {}, { engine, functions: null }), 42)
})

test('aggregates handle ranges larger than the spread-argument limit', () => {
  // Math.max(...numbers) throws RangeError past ~125k arguments, which the error
  // wrapper would have turned into a silent #NUM!.
  const big = {}
  for (let row = 1; row <= 30000; row++) for (const col of 'ABCDE') big[col + row] = row
  assert.equal(evaluate('=MAX(A1:E30000)', big), 30000)
  assert.equal(evaluate('=MIN(A1:E30000)', big), 1)
  assert.equal(evaluate('=COUNT(A1:E30000)', big), 150000)
})

test('empty text counts as blank for COUNTBLANK but not for COUNTA or ISBLANK', () => {
  // Excel is deliberately inconsistent here; verified against HyperFormula.
  const data = { A1: 1, A2: '', A3: 'x' }
  assert.equal(evaluate('=COUNTA(A1:A3)', data), 3, 'empty text still counts')
  assert.equal(evaluate('=COUNTBLANK(A1:A3)', data), 1, 'but is blank for COUNTBLANK')
  assert.equal(evaluate('=ISBLANK(A2)', data), false)
  assert.equal(evaluate('=ISBLANK(A9)', data), true, 'a cell that was never set is blank')
  assert.equal(evaluate('=COUNTA(A1:A5)', data), 3, 'unset cells are not counted')
})

test('unary functions accept both the collapsed and array forms', () => {
  // Hand-written logic may still use the array form; the engine unwraps it.
  assert.equal(defaultEngine.run({ ABS: -3 }), 3)
  assert.equal(defaultEngine.run({ ABS: [-3] }), 3)
  assert.equal(defaultEngine.run({ LEN: 'abcd' }), 4)
  assert.equal(defaultEngine.run({ LEN: ['abcd'] }), 4)
  assert.equal(defaultEngine.build({ ABS: [-3] })(), 3)
})

test('IS functions inspect errors rather than propagating them', () => {
  assert.equal(evaluate('=ISERROR(1/0)'), true)
  assert.equal(evaluate('=ISNUMBER(1/0)'), false)
  assert.equal(evaluate('=ISNA(NA())'), true)
  // ...but computing functions still propagate
  assertError('=ISEVEN(1/0)', '#DIV/0!')
  assertError('=ABS(1/0)', '#DIV/0!')
})

test('the [*] projection sugar maps a field over a collection', () => {
  const data = {
    items: [{ total: 10, qty: 2 }, { total: 20, qty: 1 }, { total: 5, qty: 9 }],
    taxRate: 0.1
  }
  assert.deepEqual(evaluate('=items[*].total', data), [10, 20, 5])
  assert.equal(evaluate('=SUM(items[*].total)', data), 35)
  assert.equal(evaluate('=MAX(items[*].qty)', data), 9)
  assert.equal(evaluate('=COUNT(items[*].total)', data), 3)
  // Composes: taxRate stays at top-level scope, outside the projection.
  assert.equal(evaluate('=SUM(items[*].total) * (1 + taxRate)', data), 38.5)
})

test('[*] projections nest for collections of collections', () => {
  const data = { rows: [{ cells: [{ v: 1 }, { v: 2 }] }, { cells: [{ v: 3 }, { v: 4 }] }] }
  assert.equal(evaluate('=SUM(rows[*].cells[*].v)', data), 10)
})
