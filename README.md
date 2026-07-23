# sheetlang

**Excel-like formulas over your own data.** Let your users write
`=IF(score >= 90, "Pass", "Fail")` or `=SUM(hours) * rate * (1 + taxRate)`, and
evaluate it safely against plain JavaScript objects — no spreadsheet grid required.

```js
import { evaluate } from 'sheetlang'

evaluate('=IF(score >= 90, "A", "B")', { score: 95 })
// 'A'

evaluate('=customer.address.city & ": " & TEXT(total, "#,##0.00")',
         { customer: { address: { city: 'Lisbon' } }, total: 1234.5 })
// 'Lisbon: 1,234.50'
```

235 built-in functions, real Excel semantics (errors are values, blanks coerce to
zero, familiar precedence), and formulas that reference your data by name or path.
Ideal for pricing rules, scoring models, computed fields, eligibility logic, and
anywhere non-technical users need to express calculations.

## Why sheetlang

- **Runs over any data shape** — nested objects, arrays, or spreadsheet-style cells.
  A formula reads `taxRate`, `customer.address.city`, `items[0].total`, or `A1:B10`.
- **Safe with untrusted input** — formulas are parsed, not `eval`'d. You control
  exactly which functions are callable, and can cap depth, array, and string sizes.
- **Formulas are just data** — every formula compiles to plain JSON you can store in
  a column, send over the wire, diff, and run later. Nothing bespoke to serialize.
- **Fast** — compile a formula once and run it over millions of rows, or drop it into
  a live `Sheet` with incremental recalculation.
- **Works everywhere** — ESM and CommonJS builds, TypeScript types included, and a
  strict-CSP mode for locked-down browsers.

## Install

```bash
npm install sheetlang
```

```js
import { evaluate } from 'sheetlang'       // ESM
const { evaluate } = require('sheetlang')  // CommonJS
```

TypeScript types ship with the package. The only runtime dependency is
`json-logic-engine`. Node 18+ or any modern bundler.

## Usage

### Evaluate a formula

```js
import { evaluate } from 'sheetlang'

evaluate('=(subtotal - discount) * (1 + taxRate)',
         { subtotal: 200, discount: 20, taxRate: 0.23 })
// 221.4
```

### Reuse a formula across many rows

`build` compiles once and returns a fast function — use it whenever the same formula
runs repeatedly.

```js
import { build } from 'sheetlang'

const price = build('=ROUND(qty * unitPrice * (1 - discount), 2)')

for (const row of orders) {
  row.total = price(row)   // millions of rows/sec
}
```

### Store formulas as data

Every formula compiles to plain JSON. Persist it in a database column, inspect it,
or transform it — then run it later with the engine.

```js
import { compile, defaultEngine } from 'sheetlang'

// Compile to portable JSON and save it anywhere.
const rule = compile('=score * weight')
// { '*': [ { val: 'score' }, { val: 'weight' } ] }

// Load it back and run it — no re-parsing needed.
defaultEngine.run(rule, { score: 8, weight: 1.5 })   // 12
```

## Writing formulas

Formulas look like Excel. The leading `=` is optional.

### Referencing your data

Read values by name, by dotted or indexed **path**, or by spreadsheet-style **cell**:

```js
evaluate('=revenue - costs',           { revenue: 100, costs: 60 })          // 40
evaluate('=customer.address.city',     { customer: { address: { city: 'Ely' } } }) // 'Ely'
evaluate('=orders[0].total',           { orders: [{ total: 42 }] })          // 42
evaluate('=[Total Sales] * 1.5',       { 'Total Sales': 1000 })              // 1500
evaluate('=A1 + B1',                   { A1: 5, B1: 10 })                    // 15
```

Use `[*]` to pull a field from every item in a collection, then aggregate it:

```js
const data = { items: [{ total: 10 }, { total: 20 }, { total: 5 }] }

evaluate('=items[*].total',      data)   // [10, 20, 5]
evaluate('=SUM(items[*].total)', data)   // 35
evaluate('=MAX(items[*].total)', data)   // 20
```

### Ranges

When your data is keyed like a grid, use ranges. They read every cell in the
rectangle, and missing cells count as blank:

```js
evaluate('=SUM(A1:B2)',   { A1: 1, B1: 2, A2: 3, B2: 4 })              // 10
evaluate('=AVERAGE(A1:A3)', { A1: 10, A2: 20, A3: 30 })                // 20
evaluate('=SUM(Sheet1!A1:A2)', { Sheet1: { A1: 10, A2: 20 } })         // 30
```

### Operators

`+  -  *  /  ^`, comparison `=  <>  <  <=  >  >=`, text join `&`, and trailing `%`.
Precedence and quirks match Excel — `-2^2` is `4`, `2^3^2` is `64`, `50%` is `0.5`.

```js
evaluate('="a" = "A"')       // true  — text compares case-insensitively
evaluate('="Total: " & 42')  // 'Total: 42'
evaluate('=blank + 5')       // 5     — a missing value counts as zero
evaluate('="41" + 1')        // 42    — numeric text converts
```

## Errors are values, not exceptions

Just like a spreadsheet, an error flows through a formula instead of throwing, and
you can catch it:

```js
evaluate('=1/0')                        // #DIV/0!
evaluate('=IFERROR(1/0, "fallback")')   // 'fallback'
evaluate('=ISERROR(SQRT(-1))')          // true
```

`#DIV/0!`, `#VALUE!`, `#REF!`, `#NAME?`, `#NUM!`, `#N/A`, `#NULL!`, and `#CIRCULAR!`
are all supported. Check a result with `isError(value)`; `value.code` is the token.

`IF`, `IFS`, `IFERROR`, `IFNA`, and `SWITCH` only evaluate the branch they take, so a
guarded expression stays safe:

```js
evaluate('=IF(qty = 0, 0, total / qty)', { total: 100, qty: 0 })   // 0, not #DIV/0!
```

## A live spreadsheet

`Sheet` connects formulas together. Cells hold a literal value or an `=` formula that
references other cells. Reads are lazy, results are cached, and editing a cell
recalculates only what depends on it. Circular references resolve to `#CIRCULAR!`
instead of hanging.

```js
import { Sheet } from 'sheetlang'

const invoice = new Sheet({
  VAT: 0.2,
  B2: 3, C2: 19.99, D2: '=B2*C2',
  B3: 1, C3: 249.00, D3: '=B3*C3',
  D6: '=SUM(D2:D3)',
  D7: '=ROUND(D6 * VAT, 2)',
  D8: '=ROUND(D6 + D7, 2)'
})

invoice.get('D8')       // 370.76
invoice.set('B3', 2)    // change quantity
invoice.get('D8')       // 669.56  — recalculated automatically
invoice.toJSON()        // { VAT: 0.2, B2: 3, ..., D8: 669.56 }
```

Recalculation stays incremental even on large sheets — editing one cell in a
10,000-cell sheet updates in well under a millisecond.

## Custom functions

Register your own functions to extend the language:

```js
import { createEngine, evaluate } from 'sheetlang'

const engine = createEngine({
  functions: {
    SLUGIFY: {
      method: ([text]) => String(text).toLowerCase().replace(/[^a-z0-9]+/g, '-'),
      deterministic: true
    }
  }
})

evaluate('=SLUGIFY(title)', { title: 'Hello World' }, { engine, functions: null })
// 'hello-world'
```

## Strict Content-Security-Policy

For speed, `build()` (and a hot `Sheet` cell) compile formulas to JavaScript, which
uses `eval`. If your page runs under a strict CSP — `script-src` without
`'unsafe-eval'` — turn on interpreted mode, which never generates code:

```js
const price = build('=ROUND(qty * unitPrice, 2)', { interpreted: true })
const sheet = new Sheet(cells, { interpreted: true })
```

`evaluate()` needs no flag — it is already interpreter-only and CSP-safe.

## TypeScript

Full type declarations are bundled. Everything is typed out of the box:

```ts
import { evaluate, build, Sheet, isError } from 'sheetlang'

const result = evaluate('=A1 + A2', { A1: 1, A2: 2 })
if (isError(result)) console.error(result.code)
```

## Function library

235 functions, all case-insensitive.

- **Math** `SUM` `PRODUCT` `SUMPRODUCT` `SUMIF` `SUMIFS` `SUMSQ` `SUMX2MY2` `SUMX2PY2` `SUMXMY2` `ABS` `SIGN` `SQRT` `SQRTPI` `POWER` `EXP` `LN` `LOG` `LOG10` `MOD` `QUOTIENT` `INT` `TRUNC` `ROUND` `ROUNDUP` `ROUNDDOWN` `MROUND` `CEILING` `CEILING.MATH` `CEILING.PRECISE` `ISO.CEILING` `FLOOR` `FLOOR.MATH` `FLOOR.PRECISE` `EVEN` `ODD` `GCD` `LCM` `FACT` `FACTDOUBLE` `COMBIN` `COMBINA` `MULTINOMIAL` `DELTA` `BASE` `DECIMAL` `ROMAN` `ARABIC` `PI` `RAND` `RANDBETWEEN`
- **Trigonometry** `SIN` `COS` `TAN` `COT` `SEC` `CSC` `ASIN` `ACOS` `ATAN` `ATAN2` `ACOT` `SINH` `COSH` `TANH` `COTH` `SECH` `CSCH` `ASINH` `ACOSH` `ATANH` `ACOTH` `RADIANS` `DEGREES`
- **Statistics** `AVERAGE` `AVERAGEA` `AVERAGEIF` `AVERAGEIFS` `COUNT` `COUNTA` `COUNTBLANK` `COUNTIF` `COUNTIFS` `COUNTUNIQUE` `MAX` `MAXA` `MAXIFS` `MIN` `MINA` `MINIFS` `MEDIAN` `MODE` `LARGE` `SMALL` `PERCENTILE` `PERCENTILE.INC` `PERCENTILE.EXC` `QUARTILE` `QUARTILE.INC` `QUARTILE.EXC` `VAR` `VAR.S` `VAR.P` `VARA` `VARP` `VARPA` `STDEV` `STDEV.S` `STDEV.P` `STDEVA` `STDEVP` `STDEVPA` `GEOMEAN` `HARMEAN` `AVEDEV` `DEVSQ` `CORREL` `RSQ` `SLOPE` `INTERCEPT` `STANDARDIZE` `FISHER` `FISHERINV` `SUBTOTAL`
- **Logical** `IF` `IFS` `IFERROR` `IFNA` `SWITCH` `AND` `OR` `XOR` `NOT` `TRUE` `FALSE`
- **Text** `CONCAT` `CONCATENATE` `TEXTJOIN` `SPLIT` `LEN` `LOWER` `UPPER` `PROPER` `TRIM` `CLEAN` `LEFT` `RIGHT` `MID` `REPT` `EXACT` `FIND` `SEARCH` `SUBSTITUTE` `REPLACE` `CHAR` `CODE` `UNICHAR` `UNICODE` `T` `N` `VALUE` `TEXT`
- **Lookup** `VLOOKUP` `HLOOKUP` `XLOOKUP` `INDEX` `MATCH` `CHOOSE` `ROWS` `COLUMNS` `TRANSPOSE`
- **Information** `ISBLANK` `ISNUMBER` `ISTEXT` `ISNONTEXT` `ISLOGICAL` `ISERROR` `ISERR` `ISNA` `ISEVEN` `ISODD` `NA` `TYPE` `ERRORTYPE`
- **Date** `TODAY` `NOW` `DATE` `TIME` `DATEVALUE` `TIMEVALUE` `YEAR` `MONTH` `DAY` `HOUR` `MINUTE` `SECOND` `WEEKDAY` `WEEKNUM` `ISOWEEKNUM` `DAYS` `DAYS360` `DATEDIF` `YEARFRAC` `EDATE` `EOMONTH` `NETWORKDAYS` `WORKDAY`
- **Financial** `PMT` `IPMT` `PPMT` `ISPMT` `CUMIPMT` `CUMPRINC` `FV` `PV` `NPV` `XNPV` `NPER` `RATE` `IRR` `MIRR` `RRI` `PDURATION` `FVSCHEDULE` `SLN` `SYD` `DB` `DDB` `EFFECT` `NOMINAL` `DOLLARDE` `DOLLARFR`

`COUNTIF`/`SUMIF`/`AVERAGEIF` and their `*IFS` forms accept Excel criteria strings
(`">10"`, `"<>x"`, `"ap*"`). `TEXT` supports common numeric codes (`#,##0.00`,
`0.0%`, `000`) and date patterns (`yyyy-mm-dd`, `d mmmm yyyy`). Dates are JavaScript
`Date` objects and behave as Excel serial numbers in arithmetic, so
`=DATE(2026,7,24) - DATE(2026,7,23)` is `1`.

## API reference

| Export | Purpose |
| --- | --- |
| `evaluate(source, data?, options?)` | Compile and run a formula in one step |
| `build(source, options?)` | Compile once, return a reusable `(data) => value` |
| `compile(source, options?)` | Turn a formula into its portable JSON form |
| `new Sheet(cells, options?)` | A live spreadsheet with incremental recalculation |
| `createEngine(options?)` | An engine with the standard library plus your functions |
| `defaultEngine` | The built-in engine; use `.run(logic, data)` to run compiled JSON |
| `isError(value)` | Whether a result is an error value |

**Options.** `functions` restricts which function names are allowed (`null` allows
any); `engine` supplies a custom engine; `interpreted` runs without code generation
for strict CSP. Engine options include `maxDepth`, `maxArrayLength`, and
`maxStringLength`.

## License

MIT
