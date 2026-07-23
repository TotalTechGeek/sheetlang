import test from 'node:test'
import assert from 'node:assert/strict'
import { evaluate, isError, functionNames } from '../src/index.js'

/**
 * The extended library, chosen by diffing against HyperFormula's 418 functions.
 *
 * Every expectation below was cross-checked against HyperFormula 3.3.0 by
 * `compare/extended.js`; INTERCEPT is the one exception, since HyperFormula has no
 * equivalent, so it is verified against exact linear data instead.
 */

const GRID = [
  [5, 12, -1000, 'a'], [3, 7, 300, 'b'], [8, 19, 400, 'a'], [1, 2, 500, 'c'],
  [9, 21, 200, 'b'], [2, 5, 150, 'a'], [7, 16, 100, 'c'], [4, 9, 50, 'a'],
  [6, 14, 25, 'b'], [10, 23, 10, 'c']
]
const DATA = {}
GRID.forEach((row, r) => row.forEach((value, c) => {
  DATA[String.fromCharCode(65 + c) + (r + 1)] = value
}))

test('multi-criteria aggregation', () => {
  assert.equal(evaluate('=COUNTIFS(A1:A10,">3",D1:D10,"a")', DATA), 3)
  assert.equal(evaluate('=SUMIFS(A1:A10,D1:D10,"a")', DATA), 19)
  assert.equal(evaluate('=SUMIFS(A1:A10,D1:D10,"a",B1:B10,">10")', DATA), 13)
  assert.equal(evaluate('=MAXIFS(A1:A10,D1:D10,"a")', DATA), 8)
  assert.equal(evaluate('=MINIFS(A1:A10,D1:D10,"b")', DATA), 3)
  assert.equal(evaluate('=COUNTUNIQUE(D1:D10)', DATA), 3)
})

test('SUBTOTAL dispatches by function number', () => {
  assert.equal(evaluate('=SUBTOTAL(9,A1:A10)', DATA), 55)
  assert.equal(evaluate('=SUBTOTAL(1,A1:A10)', DATA), 5.5)
  assert.equal(evaluate('=SUBTOTAL(4,A1:A10)', DATA), 10)
  // 101-111 mean "ignore hidden rows", which is a no-op without a UI
  assert.equal(evaluate('=SUBTOTAL(102,A1:A10)', DATA), 10)
  assert.ok(isError(evaluate('=SUBTOTAL(99,A1:A10)', DATA)))
})

test('percentile and quartile', () => {
  assert.equal(evaluate('=PERCENTILE(A1:A10,0.25)', DATA), 3.25)
  assert.equal(evaluate('=QUARTILE(A1:A10,1)', DATA), 3.25)
  assert.equal(evaluate('=PERCENTILE.EXC(A1:A10,0.5)', DATA), 5.5)
  assert.ok(isError(evaluate('=PERCENTILE(A1:A10,2)', DATA)))
})

test('descriptive statistics', () => {
  assert.equal(evaluate('=DEVSQ(A1:A10)', DATA), 82.5)
  assert.equal(evaluate('=SUMSQ(A1:A10)', DATA), 385)
  assert.equal(evaluate('=AVEDEV(A1:A10)', DATA), 2.5)
  assert.equal(evaluate('=CORREL(A1:A10,B1:B10)', DATA).toFixed(6), '0.999147')
  assert.equal(evaluate('=VAR.S(A1:A10)', DATA).toFixed(6), '9.166667')
  assert.equal(evaluate('=VAR.P(A1:A10)', DATA), 8.25)
  assert.ok(isError(evaluate('=GEOMEAN(-1, 2)')), 'GEOMEAN rejects non-positive values')
})

test('regression functions recover an exact line', () => {
  // y = 3x + 7
  const linear = {}
  for (let i = 1; i <= 5; i++) { linear['A' + i] = i; linear['B' + i] = 3 * i + 7 }
  assert.equal(evaluate('=SLOPE(B1:B5,A1:A5)', linear), 3)
  assert.equal(evaluate('=INTERCEPT(B1:B5,A1:A5)', linear), 7)
  assert.equal(evaluate('=RSQ(A1:A5,B1:B5)', linear), 1)
})

test('combinatorics and integer maths', () => {
  assert.equal(evaluate('=FACT(6)'), 720)
  assert.equal(evaluate('=FACT(0)'), 1)
  assert.equal(evaluate('=FACTDOUBLE(7)'), 105)
  assert.equal(evaluate('=COMBIN(10,3)'), 120)
  assert.equal(evaluate('=COMBINA(4,3)'), 20)
  assert.equal(evaluate('=MULTINOMIAL(2,3,4)'), 1260)
  assert.equal(evaluate('=QUOTIENT(17,5)'), 3)
  assert.equal(evaluate('=QUOTIENT(-17,5)'), -3, 'QUOTIENT truncates toward zero')
  assert.ok(isError(evaluate('=FACT(-1)')))
})

test('the CEILING/FLOOR family differs on how it treats negatives', () => {
  assert.equal(evaluate('=CEILING.MATH(-4.2)'), -4, 'default rounds toward zero')
  assert.equal(evaluate('=CEILING.MATH(-4.2,1,1)'), -5, 'a non-zero mode rounds away')
  assert.equal(evaluate('=FLOOR.MATH(-4.8)'), -5)
  assert.equal(evaluate('=FLOOR.MATH(-4.8,1,1)'), -4)
  assert.equal(evaluate('=CEILING.PRECISE(-4.2,2)'), -4)
  assert.equal(evaluate('=FLOOR.PRECISE(-4.8,2)'), -6)
})

test('number bases and roman numerals round-trip', () => {
  assert.equal(evaluate('=BASE(255,16)'), 'FF')
  assert.equal(evaluate('=BASE(255,2,12)'), '000011111111')
  assert.equal(evaluate('=DECIMAL("FF",16)'), 255)
  assert.equal(evaluate('=ROMAN(1994)'), 'MCMXCIV')
  assert.equal(evaluate('=ARABIC("MCMXCIV")'), 1994)
  assert.equal(evaluate('=ARABIC(ROMAN(3888))'), 3888)
})

test('extended trigonometry', () => {
  assert.equal(evaluate('=ASINH(1)').toFixed(6), '0.881374')
  assert.equal(evaluate('=ACOSH(2)').toFixed(6), '1.316958')
  assert.equal(evaluate('=ATANH(0.5)').toFixed(6), '0.549306')
  assert.equal(evaluate('=COT(1)').toFixed(6), '0.642093')
  assert.equal(evaluate('=SEC(1)').toFixed(6), '1.850816')
  assert.ok(isError(evaluate('=ACOSH(0)')))
  assert.ok(isError(evaluate('=CSC(0)')))
})

test('date arithmetic conventions', () => {
  assert.equal(evaluate('=DAYS360(DATE(2026,1,31),DATE(2026,3,31))'), 60)
  assert.equal(evaluate('=DATEDIF(DATE(2024,1,15),DATE(2026,7,23),"Y")'), 2)
  assert.equal(evaluate('=DATEDIF(DATE(2024,1,15),DATE(2026,7,23),"M")'), 30)
  assert.equal(evaluate('=DATEDIF(DATE(2024,1,15),DATE(2026,7,23),"MD")'), 8)
  assert.equal(evaluate('=DATEDIF(DATE(2024,1,15),DATE(2026,7,23),"YM")'), 6)
  assert.equal(evaluate('=YEARFRAC(DATE(2026,1,1),DATE(2026,7,1))'), 0.5)
  assert.equal(evaluate('=ISOWEEKNUM(DATE(2027,1,1))'), 53, 'ISO weeks belong to the year of their Thursday')
  assert.equal(evaluate('=NETWORKDAYS(DATE(2026,7,1),DATE(2026,7,31))'), 23)
  assert.equal(evaluate('=NETWORKDAYS(DATE(2026,7,1),DATE(2026,7,31),DATE(2026,7,3))'), 22)
  assert.equal(evaluate('=TIMEVALUE("14:30")'), 0.604166666666666663)
  assert.ok(isError(evaluate('=DATEDIF(DATE(2026,7,23),DATE(2024,1,15),"Y")')), 'start must precede end')
})

test('financial functions follow Excel sign conventions', () => {
  // Money leaving your pocket is negative.
  assert.equal(evaluate('=ROUND(IPMT(0.05/12,1,360,300000),2)'), -1250)
  assert.equal(evaluate('=ROUND(PPMT(0.05/12,1,360,300000),2)'), -360.46)
  // ...and the two must reconstitute the payment
  const pmt = evaluate('=ROUND(PMT(0.05/12,360,300000),2)')
  const parts = evaluate('=ROUND(IPMT(0.05/12,1,360,300000) + PPMT(0.05/12,1,360,300000),2)')
  assert.equal(parts, pmt, 'IPMT + PPMT must equal PMT')

  assert.equal(evaluate('=ROUND(NPER(0.05/12,-500,20000),6)'), 43.848271)
  assert.equal(evaluate('=ROUND(RATE(60,-500,20000),6)'), 0.014395)
  assert.equal(evaluate('=SLN(10000,1000,5)'), 1800)
  assert.equal(evaluate('=SYD(10000,1000,5,2)'), 2400)
  assert.equal(evaluate('=ROUND(EFFECT(0.05,12),6)'), 0.051162)
  assert.equal(evaluate('=ROUND(NOMINAL(EFFECT(0.05,12),12),6)'), 0.05, 'EFFECT and NOMINAL invert')
})

test('iterative solvers converge', () => {
  const flows = { C1: -1000, C2: 300, C3: 400, C4: 500, C5: 200, C6: 150 }
  assert.equal(evaluate('=ROUND(IRR(C1:C6),6)', flows), 0.186999)
  // IRR is the rate at which NPV is zero, so feeding it back must cancel out.
  // Math.abs because the residual lands on -0, which strict equality rejects.
  const npv = evaluate('=ROUND(NPV(IRR(C1:C6), C2:C6) + C1, 6)', flows)
  assert.equal(Math.abs(npv), 0)
})

test('the extended library is registered and reachable', () => {
  for (const name of ['COUNTIFS', 'SUMIFS', 'PERCENTILE', 'VAR.S', 'CEILING.MATH', 'IRR', 'ISOWEEKNUM']) {
    assert.ok(functionNames.includes(name), `${name} should be registered`)
  }
  assert.ok(functionNames.length > 230)
})

test('dotted function names parse', () => {
  assert.equal(evaluate('=VAR.S(1,2,3,4)'), evaluate('=VAR(1,2,3,4)'))
  assert.equal(evaluate('=STDEV.P(1,2,3,4)'), evaluate('=STDEVP(1,2,3,4)'))
  assert.equal(evaluate('=CEILING.MATH(4.2)'), 5)
})
