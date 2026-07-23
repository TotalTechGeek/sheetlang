/**
 * Verifies the extended function library against HyperFormula.
 *
 * Every function added after the initial 134 is exercised here. Where HyperFormula
 * has no equivalent it reports `#NAME?`, and the check records the case as
 * unverified rather than counting it as agreement.
 */

import { HyperFormula } from 'hyperformula'
import { evaluate, isError } from '../src/index.js'

const HF_CONFIG = { licenseKey: 'gpl-v3' }

//  A: sample     B: correlated  C: cashflows  D: criteria  E: rates   F: dates
const GRID = [
  [5, 12, -1000, 'a', 0.05, '2026-01-15'],
  [3, 7, 300, 'b', 0.06, '2026-03-31'],
  [8, 19, 400, 'a', 0.07, '2026-07-04'],
  [1, 2, 500, 'c', 0.02, '2026-09-30'],
  [9, 21, 200, 'b', 0.03, '2026-12-25'],
  [2, 5, 150, 'a', 0.04, '2027-02-14'],
  [7, 16, 100, 'c', 0.01, '2027-06-01'],
  [4, 9, 50, 'a', 0.05, '2027-11-11'],
  [6, 14, 25, 'b', 0.06, '2028-01-01'],
  [10, 23, 10, 'c', 0.07, '2028-05-20']
]

const cells = {}
GRID.forEach((row, r) => row.forEach((value, c) => {
  cells[String.fromCharCode(65 + c) + (r + 1)] = value
}))

const hf = HyperFormula.buildFromArray(GRID.map((row) => [...row]), HF_CONFIG)

const FORMULAS = [
  // multi-criteria
  '=COUNTIFS(A1:A10,">4")', '=COUNTIFS(A1:A10,">3",D1:D10,"a")',
  '=SUMIFS(A1:A10,D1:D10,"a")', '=SUMIFS(A1:A10,D1:D10,"a",B1:B10,">10")',
  '=MAXIFS(A1:A10,D1:D10,"a")', '=MINIFS(A1:A10,D1:D10,"b")',
  '=COUNTUNIQUE(D1:D10)', '=SUBTOTAL(9,A1:A10)', '=SUBTOTAL(1,A1:A10)',
  '=SUBTOTAL(4,A1:A10)', '=SUBTOTAL(102,A1:A10)',
  // percentile / quartile
  '=PERCENTILE(A1:A10,0.25)', '=PERCENTILE.INC(A1:A10,0.9)', '=PERCENTILE.EXC(A1:A10,0.5)',
  '=QUARTILE(A1:A10,1)', '=QUARTILE.INC(A1:A10,3)', '=QUARTILE.EXC(A1:A10,2)',
  // descriptive stats
  '=GEOMEAN(A1:A10)', '=HARMEAN(A1:A10)', '=AVEDEV(A1:A10)', '=DEVSQ(A1:A10)',
  '=SUMSQ(A1:A10)', '=CORREL(A1:A10,B1:B10)', '=RSQ(A1:A10,B1:B10)',
  '=SLOPE(B1:B10,A1:A10)', '=INTERCEPT(B1:B10,A1:A10)',
  '=STANDARDIZE(5,4,2)', '=FISHER(0.5)', '=FISHERINV(0.5)',
  '=SUMX2MY2(A1:A10,B1:B10)', '=SUMX2PY2(A1:A10,B1:B10)', '=SUMXMY2(A1:A10,B1:B10)',
  '=AVERAGEA(A1:A10)', '=MAXA(A1:A10)', '=MINA(A1:A10)',
  '=VARA(A1:A10)', '=VARPA(A1:A10)', '=STDEVA(A1:A10)', '=STDEVPA(A1:A10)',
  '=VAR.S(A1:A10)', '=VAR.P(A1:A10)', '=STDEV.S(A1:A10)', '=STDEV.P(A1:A10)',
  // maths
  '=FACT(6)', '=FACT(0)', '=FACTDOUBLE(7)', '=COMBIN(10,3)', '=COMBINA(4,3)',
  '=MULTINOMIAL(2,3,4)', '=QUOTIENT(17,5)', '=QUOTIENT(-17,5)', '=SQRTPI(4)',
  '=DELTA(3,3)', '=DELTA(3,4)',
  '=CEILING.MATH(4.2)', '=CEILING.MATH(-4.2)', '=CEILING.MATH(-4.2,1,1)',
  '=FLOOR.MATH(4.8)', '=FLOOR.MATH(-4.8)', '=FLOOR.MATH(-4.8,1,1)',
  '=CEILING.PRECISE(-4.2,2)', '=FLOOR.PRECISE(-4.8,2)', '=ISO.CEILING(-4.2,2)',
  '=BASE(255,16)', '=BASE(255,2,12)', '=DECIMAL("FF",16)', '=DECIMAL("111",2)',
  '=ROMAN(1994)', '=ROMAN(4)', '=ARABIC("MCMXCIV")', '=ARABIC("IV")',
  // trigonometry
  '=ASINH(1)', '=ACOSH(2)', '=ATANH(0.5)', '=COT(1)', '=COTH(1)',
  '=ACOT(1)', '=ACOTH(2)', '=SEC(1)', '=SECH(1)', '=CSC(1)', '=CSCH(1)',
  // text
  '=UNICHAR(9731)', '=UNICODE("A")', '=UNICODE("snow")',
  // dates
  '=DAYS360(DATE(2026,1,31),DATE(2026,3,31))',
  '=DAYS360(DATE(2026,1,31),DATE(2026,3,31),TRUE)',
  '=DATEDIF(DATE(2024,1,15),DATE(2026,7,23),"Y")',
  '=DATEDIF(DATE(2024,1,15),DATE(2026,7,23),"M")',
  '=DATEDIF(DATE(2024,1,15),DATE(2026,7,23),"D")',
  '=DATEDIF(DATE(2024,1,15),DATE(2026,7,23),"MD")',
  '=DATEDIF(DATE(2024,1,15),DATE(2026,7,23),"YM")',
  '=DATEDIF(DATE(2024,1,15),DATE(2026,7,23),"YD")',
  '=ROUND(YEARFRAC(DATE(2026,1,1),DATE(2026,7,1)),6)',
  '=ROUND(YEARFRAC(DATE(2026,1,1),DATE(2026,7,1),1),6)',
  '=ROUND(YEARFRAC(DATE(2026,1,1),DATE(2026,7,1),2),6)',
  '=ROUND(YEARFRAC(DATE(2026,1,1),DATE(2026,7,1),3),6)',
  '=ROUND(YEARFRAC(DATE(2026,1,1),DATE(2026,7,1),4),6)',
  '=WEEKNUM(DATE(2026,7,23))', '=WEEKNUM(DATE(2026,7,23),2)',
  '=ISOWEEKNUM(DATE(2026,7,23))', '=ISOWEEKNUM(DATE(2026,1,1))',
  '=ISOWEEKNUM(DATE(2027,1,1))',
  '=NETWORKDAYS(DATE(2026,7,1),DATE(2026,7,31))',
  '=YEAR(WORKDAY(DATE(2026,7,1),10))', '=DAY(WORKDAY(DATE(2026,7,1),10))',
  '=TIMEVALUE("14:30")',
  // financial
  '=ROUND(NPER(0.05/12,-500,20000),6)', '=ROUND(RATE(60,-500,20000),6)',
  '=ROUND(IPMT(0.05/12,1,360,300000),6)', '=ROUND(IPMT(0.05/12,12,360,300000),6)',
  '=ROUND(PPMT(0.05/12,1,360,300000),6)',
  '=ROUND(ISPMT(0.05/12,1,360,300000),6)',
  '=ROUND(CUMIPMT(0.05/12,360,300000,1,12,0),4)',
  '=ROUND(CUMPRINC(0.05/12,360,300000,1,12,0),4)',
  '=ROUND(IRR(C1:C6),6)', '=ROUND(MIRR(C1:C6,0.1,0.12),6)',
  '=ROUND(FVSCHEDULE(1000,E1:E3),6)',
  '=ROUND(PDURATION(0.05,1000,2000),6)', '=ROUND(RRI(10,1000,2000),6)',
  '=SLN(10000,1000,5)', '=SYD(10000,1000,5,2)',
  '=ROUND(DDB(10000,1000,5,2),6)', '=ROUND(DB(10000,1000,5,2),6)',
  '=ROUND(EFFECT(0.05,12),6)', '=ROUND(NOMINAL(0.0512,12),6)',
  '=ROUND(DOLLARDE(1.02,16),6)', '=ROUND(DOLLARFR(1.125,16),6)'
]

function normalise (value) {
  if (value === null || value === undefined) return null
  if (isError(value)) return value.code
  if (typeof value === 'object' && value !== null && value.type && value.value) return value.value
  if (value instanceof Date) return value.toISOString().slice(0, 10)
  if (Array.isArray(value)) return value.map(normalise)
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) return String(value)
    return Math.abs(value) < 1e-9 ? 0 : Number(value.toPrecision(9))
  }
  return value
}

/** Differences investigated and accepted, with the reason. */
const EXPECTED = {
  // Empty: every difference found so far was a bug on our side and has been fixed.
}

const agree = []
const differ = []
const unverified = []

for (const formula of FORMULAS) {
  const ours = normalise(evaluate(formula, cells))
  const hfSource = formula
    .replace(/\bTRUE\b(?!\()/g, 'TRUE()')
    .replace(/\bFALSE\b(?!\()/g, 'FALSE()')

  let theirs
  try {
    theirs = normalise(hf.calculateFormula(hfSource, 0))
  } catch (error) {
    unverified.push({ formula, ours, why: error.message.split('\n')[0] })
    continue
  }

  if (theirs === '#NAME?') {
    unverified.push({ formula, ours, why: 'not implemented by HyperFormula' })
    continue
  }

  if (JSON.stringify(ours) === JSON.stringify(theirs)) agree.push(formula)
  else differ.push({ formula, ours, theirs })
}

console.log(`\n\x1b[1mExtended library vs HyperFormula ${HyperFormula.version}\x1b[0m`)
console.log(`${FORMULAS.length} formulas\n`)
console.log(`  agree:       ${agree.length}`)
console.log(`  differ:      ${differ.length}`)
console.log(`  unverified:  ${unverified.length}  (no HyperFormula equivalent)`)

const explained = differ.filter((d) => EXPECTED[d.formula])
const unexplained = differ.filter((d) => !EXPECTED[d.formula])
console.log(`  ...explained: ${explained.length}`)
console.log(`  \x1b[1mUNEXPLAINED: ${unexplained.length}\x1b[0m`)

if (unexplained.length) {
  console.log('\n\x1b[1m\x1b[31mUNEXPLAINED — investigate\x1b[0m')
  const width = Math.max(...unexplained.map((d) => d.formula.length))
  for (const { formula, ours, theirs } of unexplained) {
    console.log(`  ${formula.padEnd(width)}  ours=${JSON.stringify(ours)}  hf=${JSON.stringify(theirs)}`)
  }
}

if (explained.length) {
  console.log('\n\x1b[1mKnown differences\x1b[0m')
  for (const { formula, ours, theirs } of explained) {
    console.log(`  ${formula}  ours=${JSON.stringify(ours)}  hf=${JSON.stringify(theirs)}`)
    console.log(`    \x1b[2m${EXPECTED[formula]}\x1b[0m`)
  }
}

if (unverified.length) {
  console.log('\n\x1b[1mUnverified (HyperFormula has no equivalent)\x1b[0m')
  for (const { formula, ours } of unverified) {
    console.log(`  ${formula.padEnd(46)} ours=${JSON.stringify(ours)}`)
  }
}
console.log()
