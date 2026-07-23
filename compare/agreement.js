/**
 * Differential correctness check against HyperFormula.
 *
 * Timings are the least interesting thing here. HyperFormula is a mature,
 * independently written implementation of the same semantics, so disagreement is
 * evidence of a bug in one of us — and usually in the younger one.
 */

import { HyperFormula } from 'hyperformula'
import { evaluate, isError } from '../src/index.js'

const HF_CONFIG = { licenseKey: 'gpl-v3' }

// A shared 5x6 grid, addressed as A1..E6 by both engines.
const GRID = [
  [1, 2, 3, 'apple', true],
  [10, 20, 30, 'apricot', false],
  [5, -5, 0, 'banana', true],
  [2.5, 0.5, 100, '', 1],
  [-3, 7, 1000, 'Cherry', 0],
  [1.005, 2.5, -2.5, 'date', 42]
]

const cells = {}
GRID.forEach((row, r) => row.forEach((value, c) => {
  cells[String.fromCharCode(65 + c) + (r + 1)] = value
}))

const hf = HyperFormula.buildFromArray(GRID.map((row) => [...row]), HF_CONFIG)

const FORMULAS = [
  // arithmetic and coercion
  '=A1+B1*C1', '=(A1+B1)*C1', '=A4/B4', '=A1/0', '=-C2^2', '=2^3^2', '=B2%',
  '=A6*100', '=A3-B3', '=C4/D4',
  // comparison
  '=A1=B1', '=D1="APPLE"', '=D1<D2', '=A1<D1', '=E1=TRUE', '=D4=""', '=A1<>B1',
  // text
  '=D1&"-"&D2', '=LEN(D2)', '=UPPER(D1)', '=LEFT(D2,3)', '=RIGHT(D2,3)',
  '=MID(D2,2,4)', '=TRIM("  a   b  ")', '=SUBSTITUTE(D2,"ap","AP")',
  '=FIND("ric",D2)', '=SEARCH("RIC",D2)', '=REPT(D1,2)', '=EXACT(D1,"Apple")',
  '=PROPER("hello wide world")', '=CONCATENATE(D1,D2)', '=REPLACE(D2,1,2,"XX")',
  // math
  '=ROUND(A6,2)', '=ROUND(B6,0)', '=ROUND(C6,0)', '=ROUND(-2.5,0)',
  '=ROUNDUP(1.001,2)', '=ROUNDDOWN(1.009,2)', '=INT(-1.5)', '=TRUNC(-1.5)',
  '=ABS(A5)', '=SIGN(A5)', '=SQRT(C4)', '=POWER(A2,2)', '=MOD(-3,5)', '=MOD(A2,B1)',
  '=EVEN(A1)', '=ODD(B1)', '=CEILING(4.2,0.5)', '=FLOOR(4.8,0.5)', '=GCD(24,36)',
  '=LCM(4,6)', '=EXP(0)', '=LN(A2)', '=LOG10(C4)', '=LOG(A2,10)',
  // aggregation over ranges
  '=SUM(A1:C6)', '=SUM(A1:A6)', '=AVERAGE(A1:C3)', '=COUNT(A1:E6)',
  '=COUNTA(A1:E6)', '=COUNTBLANK(A1:E6)', '=MAX(A1:C6)', '=MIN(A1:C6)',
  '=MEDIAN(A1:A6)', '=PRODUCT(A1:A3)', '=SUMPRODUCT(A1:A3,B1:B3)',
  '=STDEV(A1:A6)', '=STDEVP(A1:A6)', '=VAR(A1:A6)', '=LARGE(A1:A6,2)', '=SMALL(A1:A6,2)',
  // criteria
  '=COUNTIF(A1:A6,">2")', '=COUNTIF(D1:D6,"ap*")', '=COUNTIF(A1:A6,"<>5")',
  '=SUMIF(A1:A6,">2")', '=SUMIF(D1:D6,"ap*",A1:A6)', '=AVERAGEIF(A1:A6,">=2")',
  // logical
  '=IF(A1>0,"y","n")', '=IF(A1>100,"y","n")', '=AND(E1,A1>0)', '=OR(E2,A1>0)',
  '=NOT(E1)', '=IFERROR(A1/0,"safe")', '=IFERROR(A1/B1,"safe")',
  '=IFS(A1>100,"big",A1>0,"small")', '=SWITCH(A1,1,"one",2,"two","other")',
  // lookup
  '=VLOOKUP(5,A1:D6,4,FALSE)', '=VLOOKUP(10,A1:D6,2,FALSE)', '=VLOOKUP(4,A1:D6,2,TRUE)',
  '=HLOOKUP(2,A1:E2,2,FALSE)', '=MATCH(5,A1:A6,0)', '=INDEX(A1:D6,2,4)',
  '=INDEX(A1:A6,3)', '=CHOOSE(2,"a","b","c")', '=ROWS(A1:C6)', '=COLUMNS(A1:C6)',
  // information
  '=ISBLANK(D4)', '=ISNUMBER(A1)', '=ISTEXT(D1)', '=ISLOGICAL(E1)',
  '=ISERROR(A1/0)', '=ISEVEN(A1)', '=ISODD(A1)',
  // errors
  '=A1+D1', '=SQRT(A5)', '=VLOOKUP(999,A1:D6,2,FALSE)',
  // text formatting
  '=TEXT(1234.5,"#,##0.00")', '=TEXT(0.256,"0.0%")', '=TEXT(7,"000")',
  // dates
  '=YEAR(DATE(2026,7,23))', '=MONTH(DATE(2026,7,23))', '=DAY(DATE(2026,7,23))',
  '=DATE(2026,7,24)-DATE(2026,7,23)', '=WEEKDAY(DATE(2026,7,23))',
  // financial
  '=ROUND(PMT(0.05/12,360,300000),2)', '=ROUND(FV(0.05/12,120,-500),2)',
  '=ROUND(PV(0.05,10,-1000),2)', '=ROUND(NPV(0.1,100,200,300),2)'
]

/** Normalise both engines' outputs onto comparable ground. */
function normalise (value) {
  if (value === null || value === undefined) return null
  if (isError(value)) return value.code
  // HyperFormula returns a DetailedCellError object
  if (typeof value === 'object' && value.type && value.value) return value.value
  if (typeof value === 'object' && value.error) return String(value.error)
  if (value instanceof Date) return value.toISOString().slice(0, 10)
  if (typeof value === 'number') return Math.abs(value) < 1e-9 ? 0 : Number(value.toPrecision(10))
  return value
}

/**
 * Differences we have investigated and accepted, with the reason. Everything else
 * is treated as unexplained and worth a look.
 */
const EXPECTED = {
  '=ROUND(A6,2)': 'Excel decimal-corrects 1.005 to 1.01; HyperFormula is IEEE-honest (1.00). We follow Excel.',
  '=INT(-1.5)': 'Excel INT rounds toward -inf (-2). HyperFormula truncates (-1). We follow Excel.',
  '=MOD(-3,5)': 'Excel MOD takes the sign of the divisor (2). HyperFormula uses JS remainder (-3). We follow Excel.',
  '=COUNTBLANK(A1:E6)': 'Excel counts empty text as blank (1). HyperFormula does not (0). We follow Excel.',
  '=VLOOKUP(4,A1:D6,2,TRUE)': 'Approximate match over an UNSORTED column — undefined behaviour in Excel, so neither answer is wrong.',
  '=TEXT(1234.5,\"#,##0.00\")': 'HyperFormula does not implement this format code.',
  '=TEXT(0.256,\"0.0%\")': 'HyperFormula does not implement this format code.'
}

const agreements = []
const differences = []
const crashes = []

for (const formula of FORMULAS) {
  let ours, theirs
  try {
    ours = normalise(evaluate(formula, cells))
  } catch (error) {
    crashes.push({ formula, side: 'formulas', message: error.message.split('\n')[0] })
    continue
  }
  try {
    // HyperFormula requires TRUE()/FALSE(); Excel and we accept the bare literals.
    theirs = normalise(hf.calculateFormula(formula.replace(/\bTRUE\b(?!\()/g, 'TRUE()').replace(/\bFALSE\b(?!\()/g, 'FALSE()'), 0))
  } catch (error) {
    crashes.push({ formula, side: 'HyperFormula', message: error.message.split('\n')[0] })
    continue
  }

  const same = JSON.stringify(ours) === JSON.stringify(theirs)
  ;(same ? agreements : differences).push({ formula, ours, theirs })
}

console.log(`\n\x1b[1mDifferential check against HyperFormula ${HyperFormula.version}\x1b[0m`)
console.log(`${FORMULAS.length} formulas over a shared 5x6 grid\n`)
console.log(`  agree:     ${agreements.length}`)
console.log(`  differ:    ${differences.length}`)
console.log(`  errored:   ${crashes.length}`)

const explained = differences.filter((d) => EXPECTED[d.formula])
const unexplained = differences.filter((d) => !EXPECTED[d.formula])
console.log(`  ...of which explained: ${explained.length}`)
console.log(`  UNEXPLAINED: ${unexplained.length}`)

if (explained.length) {
  console.log('\n\x1b[1mKnown, investigated differences\x1b[0m')
  for (const { formula, ours, theirs } of explained) {
    console.log(`  ${formula}`)
    console.log(`    ours=${JSON.stringify(ours)}  hf=${JSON.stringify(theirs)}`)
    console.log(`    \x1b[2m${EXPECTED[formula]}\x1b[0m`)
  }
}

if (unexplained.length) {
  console.log('\n\x1b[1m\x1b[31mUNEXPLAINED — investigate\x1b[0m')
  for (const { formula, ours, theirs } of unexplained) {
    console.log(`  ${formula}  ours=${JSON.stringify(ours)}  hf=${JSON.stringify(theirs)}`)
  }
}

if (crashes.length) {
  console.log('\n\x1b[1mThrew\x1b[0m')
  for (const { formula, side, message } of crashes) {
    console.log(`  ${formula}  [${side}] ${message}`)
  }
}
console.log()
