/**
 * The Excel function library, expressed as json-logic-engine methods.
 *
 * Every entry here becomes a JSON Logic operation, so a compiled formula is just
 * data: `SUM(A1, 2)` is `{ SUM: [{ val: 'A1' }, 2] }`.
 *
 * Conventions:
 *  - `def(...)` wraps a plain JS function with Excel's error propagation.
 *  - `lazy(...)` declares an operation whose arguments arrive *uncompiled*, which is
 *    what IF/AND/OR need in order to short-circuit.
 */

import {
  FormulaError, isError, ERROR,
  DIV0, VALUE, REF, NUM, NA
} from './errors.js'
import {
  toNumber, toText, toBoolean, toDate, dateToSerial, serialToDate,
  flatten, findError, numericOperands, compareValues, isBlank
} from './coerce.js'
import { expandRange, resolvePath, parseA1, indexToColumn } from './refs.js'

/**
 * Wrap a plain implementation as an engine method with Excel semantics:
 * an error anywhere in the arguments short-circuits, and a thrown FormulaError
 * becomes a returned error value.
 *
 * @param {(args: any[], context: any, above: any[], engine: any) => any} handler
 * @param {{ deterministic?: boolean }} [options]
 */
function def (handler, options = {}) {
  const { deterministic = true } = options
  return {
    method: (args, context, above, engine) => {
      if (!Array.isArray(args)) args = [args]
      const failure = findError(args)
      if (failure) return failure
      try {
        return handler(args, context, above, engine)
      } catch (error) {
        if (isError(error)) return error
        if (error instanceof RangeError) return NUM
        throw error
      }
    },
    deterministic
  }
}

/**
 * A strictly one-argument operation, declared with `optimizeUnary`.
 *
 * That flag is a contract: it tells the engine the method takes its argument
 * *unwrapped*, so the optimizer stops building a one-element array per call and the
 * compiler omits the `coerceArray(...)` wrapper from the code it generates. The
 * grammar holds up the other end by emitting `{ ABS: x }` rather than `{ ABS: [x] }`.
 *
 * Only for functions that take exactly one argument — anything with an optional
 * second parameter must stay on `def`, which handles both shapes.
 *
 * @param {(value: any) => any} handler
 */
function unary (handler, options = {}) {
  const { deterministic = true } = options
  return {
    method: (value) => {
      // A range still arrives as an array here; that is the value, not a wrapper.
      if (isError(value)) return value
      try {
        return handler(value)
      } catch (error) {
        if (isError(error)) return error
        if (error instanceof RangeError) return NUM
        throw error
      }
    },
    optimizeUnary: true,
    deterministic
  }
}

/**
 * A unary operation that *inspects* its argument instead of computing with it, so
 * an error argument is data rather than something to propagate: `ISERROR(1/0)` is
 * TRUE, not `#DIV/0!`.
 *
 * @param {(value: any) => any} handler
 */
function inspect (handler) {
  return { method: handler, optimizeUnary: true, deterministic: true }
}

/**
 * Numeric aggregation over a single fused pass.
 *
 * The generic `def` path would walk a range three times — once to scan for errors,
 * once to flatten, once to collect numbers — and allocate a copy in the middle. For
 * `SUM(A1:E200)` that tripled the per-cell cost, so these functions gather numbers
 * and detect errors together, into one array.
 *
 * @param {(numbers: number[]) => any} handler
 */
function aggregate (handler) {
  return {
    method: (args) => {
      if (!Array.isArray(args)) args = [args]
      const numbers = []
      try {
        for (const arg of args) {
          if (Array.isArray(arg)) {
            const failure = collectNumbers(arg, numbers)
            if (failure) return failure
          } else if (isError(arg)) {
            return arg
          } else if (!isBlank(arg) || arg === false) {
            // A blank scalar contributes nothing; that is why AVERAGE of one is #DIV/0!.
            numbers.push(toNumber(arg))
          }
        }
        return handler(numbers)
      } catch (error) {
        if (isError(error)) return error
        if (error instanceof RangeError) return NUM
        throw error
      }
    },
    deterministic: true
  }
}

/**
 * Collect the numbers out of a (possibly nested) range, skipping the text, blanks
 * and booleans that Excel ignores inside one.
 * @returns the first error encountered, or undefined
 */
function collectNumbers (values, out) {
  for (const value of values) {
    if (typeof value === 'number') out.push(value)
    else if (Array.isArray(value)) {
      const failure = collectNumbers(value, out)
      if (failure) return failure
    } else if (isError(value)) return value
    else if (value instanceof Date) out.push(dateToSerial(value))
  }
  return undefined
}

/**
 * Declare a lazy operation. `run(logic)` evaluates one argument on demand, so
 * untaken branches cost nothing and never surface their errors.
 */
function lazy (handler) {
  return {
    lazy: true,
    deterministic: false,
    method: (args, context, above, engine) => {
      if (!Array.isArray(args)) args = [args]
      const run = (logic) => {
        const result = engine.run(logic, context, { above })
        return result === undefined ? null : result
      }
      try {
        return handler(args, run)
      } catch (error) {
        if (isError(error)) return error
        throw error
      }
    }
  }
}

/** Guard for numeric results Excel refuses to produce. */
function finite (n) {
  if (Number.isNaN(n)) return NUM
  if (!Number.isFinite(n)) return NUM
  return n
}

/** Excel rounds half away from zero, unlike Math.round which rounds half up. */
function roundHalfAway (value, digits = 0) {
  const factor = Math.pow(10, digits)
  const scaled = value * factor
  // Nudge past binary representation error (e.g. 1.005 * 100 === 100.49999...).
  const corrected = Number(scaled.toPrecision(15))
  const rounded = corrected < 0
    ? -Math.round(-corrected)
    : Math.round(corrected)
  return rounded / factor
}

// ---------------------------------------------------------------------------
// Criteria matching, shared by COUNTIF / SUMIF / AVERAGEIF
// ---------------------------------------------------------------------------

const CRITERIA_PATTERN = /^(<=|>=|<>|=|<|>)?(.*)$/s

/** Turn `">5"`, `"<>x"` or `"app*"` into a predicate. */
function criteriaPredicate (criteria) {
  if (typeof criteria !== 'string') {
    return (value) => compareValues(value, criteria) === 0
  }
  const [, operator = '', rest] = CRITERIA_PATTERN.exec(criteria)
  let operand = rest
  const asNumber = Number(operand)
  if (operand.trim() !== '' && !Number.isNaN(asNumber)) operand = asNumber
  else if (/^true$/i.test(operand)) operand = true
  else if (/^false$/i.test(operand)) operand = false

  switch (operator) {
    case '<': return (value) => compareValues(value, operand) < 0
    case '<=': return (value) => compareValues(value, operand) <= 0
    case '>': return (value) => compareValues(value, operand) > 0
    case '>=': return (value) => compareValues(value, operand) >= 0
    case '<>': return negate(matcher(operand))
    default: return matcher(operand)
  }
}

function negate (predicate) {
  return (value) => !predicate(value)
}

const escapeRegex = (text) => text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')

/** Excel wildcards: `*` any run, `?` one character, `~` escapes either. */
function wildcardToRegex (pattern) {
  let source = ''
  for (let i = 0; i < pattern.length; i++) {
    const char = pattern[i]
    if (char === '~' && '*?~'.includes(pattern[i + 1])) source += escapeRegex(pattern[++i])
    else if (char === '*') source += '.*'
    else if (char === '?') source += '.'
    else source += escapeRegex(char)
  }
  return new RegExp(`^${source}$`, 'i')
}

/** Equality that understands `*` and `?` wildcards in text criteria. */
function matcher (operand) {
  if (typeof operand === 'string' && /[*?]/.test(operand)) {
    const regex = wildcardToRegex(operand)
    return (value) => typeof value === 'string' && regex.test(value)
  }
  return (value) => {
    if (isBlank(value) && isBlank(operand)) return true
    if (isBlank(value) || isError(value)) return false
    return compareValues(value, operand) === 0
  }
}

// ---------------------------------------------------------------------------
// Lookup helpers
// ---------------------------------------------------------------------------

/** Normalise a range/array argument into a 2D grid. */
function toGrid (value) {
  if (!Array.isArray(value)) return [[value]]
  if (value.length === 0) return [[]]
  for (let i = 0; i < value.length; i++) {
    if (!Array.isArray(value[i])) return [value]
  }
  return value
}

/** A range argument as a flat vector, for the single-row/column lookup family. */
function toVector (value) {
  return flatten([value])
}

// ---------------------------------------------------------------------------
// Number formatting for TEXT()
// ---------------------------------------------------------------------------

const DATE_FORMAT_TOKENS = /(yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|AM\/PM)/g
const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July',
  'August', 'September', 'October', 'November', 'December']
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']

function looksLikeDateFormat (format) {
  return /[dy]/i.test(format) && !/[#0]/.test(format)
}

function formatDate (date, format) {
  const pad = (n, width = 2) => String(n).padStart(width, '0')
  const hours12 = date.getHours() % 12 === 0 ? 12 : date.getHours() % 12
  const usesMeridiem = /AM\/PM/i.test(format)
  return format.replace(DATE_FORMAT_TOKENS, (token) => {
    switch (token.toLowerCase()) {
      case 'yyyy': return String(date.getFullYear())
      case 'yy': return pad(date.getFullYear() % 100)
      case 'mmmm': return MONTHS[date.getMonth()]
      case 'mmm': return MONTHS[date.getMonth()].slice(0, 3)
      case 'mm': return pad(date.getMonth() + 1)
      case 'm': return String(date.getMonth() + 1)
      case 'dddd': return DAYS[date.getDay()]
      case 'ddd': return DAYS[date.getDay()].slice(0, 3)
      case 'dd': return pad(date.getDate())
      case 'd': return String(date.getDate())
      case 'hh': return pad(usesMeridiem ? hours12 : date.getHours())
      case 'h': return String(usesMeridiem ? hours12 : date.getHours())
      case 'ss': return pad(date.getSeconds())
      case 's': return String(date.getSeconds())
      case 'am/pm': return date.getHours() < 12 ? 'AM' : 'PM'
      default: return token
    }
  }).replace(/mm(?=:)|(?<=:)mm/g, pad(date.getMinutes()))
}

/**
 * A useful subset of Excel's numeric format codes: `0`, `#`, `.`, thousands `,`
 * and a trailing `%`, plus surrounding literal text.
 */
function formatNumber (value, format) {
  let n = value
  let suffix = ''
  let working = format

  if (working.includes('%')) {
    n *= 100
    working = working.replace(/%/g, '')
    suffix = '%'
  }

  const match = /([#0,]*)(?:\.([#0]*))?/.exec(working)
  const intPattern = match[1] || ''
  const decPattern = match[2] || ''
  const grouped = intPattern.includes(',')

  const decimals = decPattern.length
  const fixed = Math.abs(roundHalfAway(n, decimals)).toFixed(decimals)
  let [intPart, decPart = ''] = fixed.split('.')

  const minIntDigits = (intPattern.replace(/,/g, '').match(/0/g) || []).length
  if (intPart.length < minIntDigits) intPart = intPart.padStart(minIntDigits, '0')
  if (minIntDigits === 0 && intPart === '0' && decimals > 0) intPart = ''
  if (grouped) intPart = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',')

  // Trim digits that a `#` placeholder makes optional.
  if (decPart) {
    const optional = (decPattern.match(/#/g) || []).length
    if (optional) {
      const required = decPattern.length - optional
      decPart = decPart.replace(/0+$/, (zeros) =>
        zeros.slice(0, Math.max(0, required - (decPart.length - zeros.length))))
    }
  }

  const sign = n < 0 ? '-' : ''
  const body = decPart ? `${intPart}.${decPart}` : intPart
  const prefix = working.slice(0, match.index)
  const trailer = working.slice(match.index + match[0].length)
  return `${prefix}${sign}${body}${suffix}${trailer}`
}

// ---------------------------------------------------------------------------
// The library
// ---------------------------------------------------------------------------

export const functions = {
  // -- Reference expansion ---------------------------------------------------

  /**
   * `A1:B3` compiles to `{ RANGE: ['A1', 'B3'] }`, optionally with a third
   * argument naming the scope to read from (`Sheet1!A1:B3`).
   * Resolves to a row-major grid of values, with missing cells as null.
   */
  RANGE: def(([start, end, scope], context) => {
    const root = scope ? resolvePath(context, flatten([scope])) : context
    if (root === null || root === undefined) return REF

    // `names` is a shared, cached grid — copy values out rather than mapping over it.
    const names = expandRange(toText(start), toText(end))
    const grid = new Array(names.length)
    for (let r = 0; r < names.length; r++) {
      const row = names[r]
      const values = new Array(row.length)
      for (let c = 0; c < row.length; c++) {
        const value = root[row[c]]
        values[c] = value === undefined ? null : value
      }
      grid[r] = values
    }
    return grid
  }, { deterministic: false }),

  // -- Operators (Excel semantics, native JSON Logic names) ------------------

  '+': def((args) => {
    if (args.length === 1) return finite(toNumber(args[0]))
    let sum = 0
    for (let i = 0; i < args.length; i++) sum += toNumber(args[i])
    return finite(sum)
  }),
  '-': def((args) => {
    if (args.length === 1) return finite(-toNumber(args[0]))
    let total = toNumber(args[0])
    for (let i = 1; i < args.length; i++) total -= toNumber(args[i])
    return finite(total)
  }),
  '*': def((args) => {
    let product = 1
    for (let i = 0; i < args.length; i++) product *= toNumber(args[i])
    return finite(product)
  }),
  '/': def((args) => {
    let acc = toNumber(args[0])
    for (let i = 1; i < args.length; i++) {
      const divisor = toNumber(args[i])
      if (divisor === 0) return DIV0
      acc /= divisor
    }
    return finite(acc)
  }),

  '==': def(([a, b]) => compareValues(a, b) === 0),
  '!=': def(([a, b]) => compareValues(a, b) !== 0),
  '<': def(([a, b]) => compareValues(a, b) < 0),
  '<=': def(([a, b]) => compareValues(a, b) <= 0),
  '>': def(([a, b]) => compareValues(a, b) > 0),
  '>=': def(([a, b]) => compareValues(a, b) >= 0),

  // -- Logical ---------------------------------------------------------------

  /** Lazy so that `IF(B1=0, 0, A1/B1)` never evaluates the guarded branch. */
  IF: lazy((args, run) => {
    const test = run(args[0])
    if (isError(test)) return test
    if (toBoolean(test)) return args.length > 1 ? run(args[1]) : true
    return args.length > 2 ? run(args[2]) : false
  }),

  IFS: lazy((args, run) => {
    for (let i = 0; i + 1 < args.length; i += 2) {
      const test = run(args[i])
      if (isError(test)) return test
      if (toBoolean(test)) return run(args[i + 1])
    }
    return NA
  }),

  IFERROR: lazy((args, run) => {
    const value = run(args[0])
    return isError(value) ? run(args[1]) : value
  }),

  IFNA: lazy((args, run) => {
    const value = run(args[0])
    return isError(value) && value.code === '#N/A' ? run(args[1]) : value
  }),

  SWITCH: lazy((args, run) => {
    const subject = run(args[0])
    if (isError(subject)) return subject
    let i = 1
    for (; i + 1 < args.length; i += 2) {
      const candidate = run(args[i])
      if (isError(candidate)) return candidate
      if (compareValues(subject, candidate) === 0) return run(args[i + 1])
    }
    // An odd trailing argument is the default.
    return i < args.length ? run(args[i]) : NA
  }),

  AND: def((args) => {
    const values = flatten(args)
    let considered = 0
    for (let i = 0; i < values.length; i++) {
      const value = values[i]
      if (isBlank(value) && value !== false) continue
      considered++
      if (!toBoolean(value)) return false
    }
    return considered ? true : VALUE
  }),
  OR: def((args) => {
    const values = flatten(args)
    let considered = 0
    for (let i = 0; i < values.length; i++) {
      const value = values[i]
      if (isBlank(value) && value !== false) continue
      considered++
      if (toBoolean(value)) return true
    }
    return considered ? false : VALUE
  }),
  XOR: def((args) => {
    const values = flatten(args)
    let truths = 0
    for (let i = 0; i < values.length; i++) if (toBoolean(values[i])) truths++
    return truths % 2 === 1
  }),
  NOT: unary((value) => !toBoolean(value)),
  TRUE: def(() => true),
  FALSE: def(() => false),

  // -- Math ------------------------------------------------------------------

  SUM: aggregate((numbers) => {
    let total = 0
    for (let i = 0; i < numbers.length; i++) total += numbers[i]
    return finite(total)
  }),
  PRODUCT: aggregate((numbers) => {
    if (!numbers.length) return 0
    let product = 1
    for (let i = 0; i < numbers.length; i++) product *= numbers[i]
    return finite(product)
  }),
  SUMPRODUCT: def((args) => {
    const vectors = new Array(args.length)
    for (let i = 0; i < args.length; i++) vectors[i] = flatten([args[i]])
    const length = vectors.length ? vectors[0].length : 0
    for (let i = 0; i < vectors.length; i++) {
      if (vectors[i].length !== length) return VALUE
    }
    let total = 0
    for (let i = 0; i < length; i++) {
      let product = 1
      for (const vector of vectors) {
        const value = vector[i]
        product *= typeof value === 'number' ? value : isBlank(value) ? 0 : toNumber(value)
      }
      total += product
    }
    return finite(total)
  }),
  SUMIF: def(([range, criteria, sumRange]) => {
    const source = flatten([range])
    const target = sumRange === undefined ? source : flatten([sumRange])
    const matches = criteriaPredicate(criteria)
    let total = 0
    for (let i = 0; i < source.length; i++) {
      if (!matches(source[i])) continue
      const summand = target[i]
      if (typeof summand === 'number') total += summand
    }
    return finite(total)
  }),
  ABS: unary((n) => Math.abs(toNumber(n))),
  SIGN: unary((n) => Math.sign(toNumber(n))),
  SQRT: unary((n) => {
    const value = toNumber(n)
    if (value < 0) return NUM
    return Math.sqrt(value)
  }),
  POWER: def(([base, exponent]) => finite(Math.pow(toNumber(base), toNumber(exponent)))),
  EXP: unary((n) => finite(Math.exp(toNumber(n)))),
  LN: unary((n) => {
    const value = toNumber(n)
    if (value <= 0) return NUM
    return Math.log(value)
  }),
  LOG: def(([n, base]) => {
    const value = toNumber(n)
    const b = base === undefined ? 10 : toNumber(base)
    if (value <= 0 || b <= 0 || b === 1) return NUM
    return Math.log(value) / Math.log(b)
  }),
  LOG10: unary((n) => {
    const value = toNumber(n)
    if (value <= 0) return NUM
    return Math.log10(value)
  }),
  MOD: def(([n, divisor]) => {
    const d = toNumber(divisor)
    if (d === 0) return DIV0
    const value = toNumber(n)
    // Excel's MOD takes the sign of the divisor, unlike JS's %.
    return value - d * Math.floor(value / d)
  }),
  INT: unary((n) => Math.floor(toNumber(n))),
  TRUNC: def(([n, digits]) => {
    const factor = Math.pow(10, digits === undefined ? 0 : Math.trunc(toNumber(digits)))
    return Math.trunc(toNumber(n) * factor) / factor
  }),
  ROUND: def(([n, digits]) => roundHalfAway(toNumber(n), digits === undefined ? 0 : Math.trunc(toNumber(digits)))),
  ROUNDUP: def(([n, digits]) => {
    const places = digits === undefined ? 0 : Math.trunc(toNumber(digits))
    const factor = Math.pow(10, places)
    const value = toNumber(n) * factor
    return (value < 0 ? -Math.ceil(-value) : Math.ceil(value)) / factor
  }),
  ROUNDDOWN: def(([n, digits]) => {
    const places = digits === undefined ? 0 : Math.trunc(toNumber(digits))
    const factor = Math.pow(10, places)
    const value = toNumber(n) * factor
    return (value < 0 ? -Math.floor(-value) : Math.floor(value)) / factor
  }),
  MROUND: def(([n, multiple]) => {
    const m = toNumber(multiple)
    if (m === 0) return 0
    const value = toNumber(n)
    if (Math.sign(value) !== Math.sign(m) && value !== 0) return NUM
    return roundHalfAway(value / m) * m
  }),
  CEILING: def(([n, significance]) => {
    const s = significance === undefined ? 1 : toNumber(significance)
    if (s === 0) return 0
    const value = toNumber(n)
    if (value > 0 && s < 0) return NUM
    return Math.ceil(value / s) * s
  }),
  FLOOR: def(([n, significance]) => {
    const s = significance === undefined ? 1 : toNumber(significance)
    if (s === 0) return DIV0
    const value = toNumber(n)
    if (value > 0 && s < 0) return NUM
    return Math.floor(value / s) * s
  }),
  EVEN: unary((n) => {
    const value = toNumber(n)
    const rounded = Math.ceil(Math.abs(value) / 2) * 2
    return value < 0 ? -rounded : rounded
  }),
  ODD: unary((n) => {
    const value = toNumber(n)
    let rounded = Math.ceil(Math.abs(value))
    if (rounded % 2 === 0) rounded += 1
    if (rounded === 0) rounded = 1
    return value < 0 ? -rounded : rounded
  }),
  GCD: def((args) => {
    const numbers = numericOperands(args)
    let result = 0
    for (let i = 0; i < numbers.length; i++) {
      let a = result
      let b = Math.abs(Math.trunc(numbers[i]))
      while (b) { const next = a % b; a = b; b = next }
      result = a
    }
    return result
  }),
  LCM: def((args) => {
    const numbers = numericOperands(args)
    let result = numbers.length ? 1 : 0
    for (let i = 0; i < numbers.length; i++) {
      const value = Math.abs(Math.trunc(numbers[i]))
      if (!result || !value) { result = 0; continue }
      let a = result
      let b = value
      while (b) { const next = a % b; a = b; b = next }
      result = (result * value) / a
    }
    return result
  }),
  PI: def(() => Math.PI),
  RADIANS: unary((n) => (toNumber(n) * Math.PI) / 180),
  DEGREES: unary((n) => (toNumber(n) * 180) / Math.PI),
  SIN: unary((n) => Math.sin(toNumber(n))),
  COS: unary((n) => Math.cos(toNumber(n))),
  TAN: unary((n) => Math.tan(toNumber(n))),
  ASIN: unary((n) => finite(Math.asin(toNumber(n)))),
  ACOS: unary((n) => finite(Math.acos(toNumber(n)))),
  ATAN: unary((n) => Math.atan(toNumber(n))),
  ATAN2: def(([x, y]) => Math.atan2(toNumber(y), toNumber(x))),
  SINH: unary((n) => Math.sinh(toNumber(n))),
  COSH: unary((n) => Math.cosh(toNumber(n))),
  TANH: unary((n) => Math.tanh(toNumber(n))),
  RAND: def(() => Math.random(), { deterministic: false }),
  RANDBETWEEN: def(([low, high]) => {
    const min = Math.ceil(toNumber(low))
    const max = Math.floor(toNumber(high))
    return min + Math.floor(Math.random() * (max - min + 1))
  }, { deterministic: false }),

  // -- Statistics ------------------------------------------------------------

  AVERAGE: aggregate((numbers) => {
    if (!numbers.length) return DIV0
    let total = 0
    for (let i = 0; i < numbers.length; i++) total += numbers[i]
    return total / numbers.length
  }),
  AVERAGEIF: def(([range, criteria, averageRange]) => {
    const source = flatten([range])
    const target = averageRange === undefined ? source : flatten([averageRange])
    const matches = criteriaPredicate(criteria)
    let total = 0
    let count = 0
    for (let i = 0; i < source.length; i++) {
      if (matches(source[i]) && typeof target[i] === 'number') { total += target[i]; count++ }
    }
    if (!count) return DIV0
    return total / count
  }),
  COUNT: aggregate((numbers) => numbers.length),
  // Excel is inconsistent here, deliberately: empty text is *not* empty for COUNTA
  // (a formula returning "" still counts) but *is* empty for COUNTBLANK.
  COUNTA: def((args) => {
    const values = flatten(args)
    let count = 0
    for (let i = 0; i < values.length; i++) {
      if (values[i] !== null && values[i] !== undefined) count++
    }
    return count
  }),
  COUNTBLANK: def((args) => {
    const values = flatten(args)
    let count = 0
    for (let i = 0; i < values.length; i++) if (isBlank(values[i])) count++
    return count
  }),
  COUNTIF: def(([range, criteria]) => {
    const matches = criteriaPredicate(criteria)
    const values = flatten([range])
    let count = 0
    for (let i = 0; i < values.length; i++) if (matches(values[i])) count++
    return count
  }),
  MAX: aggregate((numbers) => {
    let best = -Infinity
    for (const n of numbers) if (n > best) best = n
    return numbers.length ? best : 0
  }),
  MIN: aggregate((numbers) => {
    let best = Infinity
    for (const n of numbers) if (n < best) best = n
    return numbers.length ? best : 0
  }),
  MEDIAN: aggregate((numbers) => {
    numbers.sort((a, b) => a - b)
    if (!numbers.length) return NUM
    const mid = numbers.length >> 1
    return numbers.length % 2 ? numbers[mid] : (numbers[mid - 1] + numbers[mid]) / 2
  }),
  MODE: aggregate((numbers) => {
    const counts = new Map()
    let best = null
    let bestCount = 1
    for (const n of numbers) {
      const count = (counts.get(n) || 0) + 1
      counts.set(n, count)
      if (count > bestCount) { best = n; bestCount = count }
    }
    return best === null ? NA : best
  }),
  LARGE: def(([range, k]) => {
    const numbers = numericOperands([range]).sort((a, b) => b - a)
    const index = Math.trunc(toNumber(k)) - 1
    if (index < 0 || index >= numbers.length) return NUM
    return numbers[index]
  }),
  SMALL: def(([range, k]) => {
    const numbers = numericOperands([range]).sort((a, b) => a - b)
    const index = Math.trunc(toNumber(k)) - 1
    if (index < 0 || index >= numbers.length) return NUM
    return numbers[index]
  }),
  VAR: aggregate((numbers) => variance(numbers, true)),
  VARP: aggregate((numbers) => variance(numbers, false)),
  STDEV: aggregate((numbers) => {
    const result = variance(numbers, true)
    return isError(result) ? result : Math.sqrt(result)
  }),
  STDEVP: aggregate((numbers) => {
    const result = variance(numbers, false)
    return isError(result) ? result : Math.sqrt(result)
  }),

  // -- Text ------------------------------------------------------------------

  CONCAT: def((args) => concatenate(args)),
  CONCATENATE: def((args) => concatenate(args)),
  TEXTJOIN: def(([delimiter, ignoreEmpty, ...rest]) => {
    const skipEmpty = toBoolean(ignoreEmpty)
    const separator = toText(delimiter)
    const values = flatten(rest)
    let result = ''
    let first = true
    for (let i = 0; i < values.length; i++) {
      if (skipEmpty && isBlank(values[i])) continue
      if (!first) result += separator
      result += toText(values[i])
      first = false
    }
    return result
  }),
  LEN: unary((value) => toText(value).length),
  LOWER: unary((value) => toText(value).toLowerCase()),
  UPPER: unary((value) => toText(value).toUpperCase()),
  PROPER: unary((value) => toText(value).replace(/(^|[^A-Za-z'])([a-z])/g,
    (_, prefix, letter) => prefix + letter.toUpperCase())),
  TRIM: unary((value) => toText(value).replace(/\s+/g, ' ').trim()),
  CLEAN: unary((value) => toText(value).replace(/[\x00-\x1F\x7F]/g, '')),
  LEFT: def(([value, count]) => {
    const n = count === undefined ? 1 : Math.trunc(toNumber(count))
    if (n < 0) return VALUE
    return toText(value).slice(0, n)
  }),
  RIGHT: def(([value, count]) => {
    const n = count === undefined ? 1 : Math.trunc(toNumber(count))
    if (n < 0) return VALUE
    return n === 0 ? '' : toText(value).slice(-n)
  }),
  MID: def(([value, start, count]) => {
    const from = Math.trunc(toNumber(start))
    const length = Math.trunc(toNumber(count))
    if (from < 1 || length < 0) return VALUE
    return toText(value).substr(from - 1, length)
  }),
  REPT: def(([value, count]) => {
    const times = Math.trunc(toNumber(count))
    if (times < 0) return VALUE
    return toText(value).repeat(times)
  }),
  EXACT: def(([a, b]) => toText(a) === toText(b)),
  FIND: def(([needle, haystack, start]) => {
    const from = start === undefined ? 1 : Math.trunc(toNumber(start))
    const index = toText(haystack).indexOf(toText(needle), from - 1)
    return index === -1 ? VALUE : index + 1
  }),
  SEARCH: def(([needle, haystack, start]) => {
    const from = start === undefined ? 1 : Math.trunc(toNumber(start))
    // SEARCH is case-insensitive and honours wildcards, unlike FIND.
    const regex = new RegExp(wildcardToRegex(toText(needle)).source.replace(/^\^|\$$/g, ''), 'i')
    const text = toText(haystack)
    const index = text.slice(from - 1).search(regex)
    return index === -1 ? VALUE : index + from
  }),
  SUBSTITUTE: def(([value, from, to, occurrence]) => {
    const text = toText(value)
    const needle = toText(from)
    const replacement = toText(to)
    if (needle === '') return text
    if (occurrence === undefined) return text.split(needle).join(replacement)
    const nth = Math.trunc(toNumber(occurrence))
    if (nth < 1) return VALUE
    let seen = 0
    let index = text.indexOf(needle)
    while (index !== -1) {
      if (++seen === nth) {
        return text.slice(0, index) + replacement + text.slice(index + needle.length)
      }
      index = text.indexOf(needle, index + needle.length)
    }
    return text
  }),
  REPLACE: def(([value, start, length, replacement]) => {
    const text = toText(value)
    const from = Math.trunc(toNumber(start))
    const count = Math.trunc(toNumber(length))
    if (from < 1 || count < 0) return VALUE
    return text.slice(0, from - 1) + toText(replacement) + text.slice(from - 1 + count)
  }),
  CHAR: unary((code) => String.fromCharCode(Math.trunc(toNumber(code)))),
  CODE: unary((value) => {
    const text = toText(value)
    return text.length ? text.charCodeAt(0) : VALUE
  }),
  T: unary((value) => (typeof value === 'string' ? value : '')),
  N: unary((value) => {
    if (typeof value === 'number') return value
    if (typeof value === 'boolean') return value ? 1 : 0
    if (value instanceof Date) return dateToSerial(value)
    return 0
  }),
  VALUE: unary((value) => toNumber(value)),
  TEXT: def(([value, format]) => {
    const pattern = toText(format)
    if (looksLikeDateFormat(pattern)) return formatDate(toDate(value), pattern)
    if (/[#0]/.test(pattern)) return formatNumber(toNumber(value), pattern)
    return toText(value)
  }),

  // -- Lookup ----------------------------------------------------------------

  CHOOSE: def(([index, ...options]) => {
    const i = Math.trunc(toNumber(index))
    if (i < 1 || i > options.length) return VALUE
    return options[i - 1]
  }),
  INDEX: def(([array, rowNum, colNum]) => {
    const grid = toGrid(array)
    const row = rowNum === undefined ? 0 : Math.trunc(toNumber(rowNum))
    const col = colNum === undefined ? 0 : Math.trunc(toNumber(colNum))
    if (row < 0 || col < 0) return VALUE

    // A single-row or single-column range indexes linearly, as Excel allows.
    if (col === 0 && grid.length === 1) return pick(grid[0], row)
    if (col === 0 && row > 0 && isSingleColumn(grid)) return pick(firstColumn(grid), row)
    if (row === 0 && col > 0) {
      const column = new Array(grid.length)
      for (let i = 0; i < grid.length; i++) column[i] = grid[i][col - 1]
      return column
    }
    if (col === 0) return grid[row - 1] ?? REF
    const target = grid[row - 1]?.[col - 1]
    return target === undefined ? REF : target
  }),
  MATCH: def(([lookup, array, matchType]) => {
    const vector = toVector(array)
    const type = matchType === undefined ? 1 : Math.trunc(toNumber(matchType))
    if (type === 0) {
      const matches = matcher(lookup)
      for (let i = 0; i < vector.length; i++) {
        if (matches(vector[i])) return i + 1
      }
      return NA
    }
    // Ascending (1) or descending (-1): the last value that does not overshoot.
    let best = -1
    for (let i = 0; i < vector.length; i++) {
      if (isBlank(vector[i])) continue
      const comparison = compareValues(vector[i], lookup)
      if (type === 1 ? comparison <= 0 : comparison >= 0) best = i
      else break
    }
    return best === -1 ? NA : best + 1
  }),
  VLOOKUP: def(([lookup, table, colIndex, approximate]) => {
    const grid = toGrid(table)
    const col = Math.trunc(toNumber(colIndex))
    if (col < 1) return VALUE
    if (grid[0] && col > grid[0].length) return REF
    const useApproximate = approximate === undefined ? true : toBoolean(approximate)
    const rowIndex = findRow(firstColumn(grid), lookup, useApproximate)
    if (rowIndex === -1) return NA
    return grid[rowIndex][col - 1] ?? null
  }),
  HLOOKUP: def(([lookup, table, rowIndex, approximate]) => {
    const grid = toGrid(table)
    const row = Math.trunc(toNumber(rowIndex))
    if (row < 1) return VALUE
    if (row > grid.length) return REF
    const useApproximate = approximate === undefined ? true : toBoolean(approximate)
    const colIndex = findRow(grid[0] ?? [], lookup, useApproximate)
    if (colIndex === -1) return NA
    return grid[row - 1][colIndex] ?? null
  }),
  XLOOKUP: def(([lookup, lookupArray, returnArray, ifNotFound]) => {
    const haystack = toVector(lookupArray)
    const results = toVector(returnArray)
    const matches = matcher(lookup)
    for (let i = 0; i < haystack.length; i++) {
      if (matches(haystack[i])) return results[i] ?? null
    }
    return ifNotFound === undefined ? NA : ifNotFound
  }),
  ROWS: unary((array) => toGrid(array).length),
  COLUMNS: unary((array) => toGrid(array)[0]?.length ?? 0),
  TRANSPOSE: unary((array) => {
    const grid = toGrid(array)
    const width = grid.length && grid[0] ? grid[0].length : 0
    const result = new Array(width)
    for (let col = 0; col < width; col++) {
      const line = new Array(grid.length)
      for (let row = 0; row < grid.length; row++) line[row] = grid[row][col] ?? null
      result[col] = line
    }
    return result
  }),

  // -- Information -----------------------------------------------------------

  // A cell holding empty text is not blank — matching Excel, and COUNTA above.
  ISBLANK: inspect((value) => value === null || value === undefined),
  ISNUMBER: inspect((value) => typeof value === 'number' || value instanceof Date),
  ISTEXT: inspect((value) => typeof value === 'string'),
  ISNONTEXT: inspect((value) => typeof value !== 'string'),
  ISLOGICAL: inspect((value) => typeof value === 'boolean'),
  ISERROR: inspect((value) => isError(value)),
  ISERR: inspect((value) => isError(value) && value.code !== '#N/A'),
  ISNA: inspect((value) => isError(value) && value.code === '#N/A'),
  ISEVEN: unary((value) => Math.trunc(toNumber(value)) % 2 === 0),
  ISODD: unary((value) => Math.abs(Math.trunc(toNumber(value))) % 2 === 1),
  NA: { method: () => NA, deterministic: true },
  /** What a literal `#DIV/0!` in a formula compiles to. */
  ERRORVALUE: inspect((code) => ERROR[code] ?? NA),
  ERRORTYPE: inspect((value) => (isError(value) ? Object.keys(ERROR).indexOf(value.code) + 1 : NA)),
  TYPE: inspect((value) => {
    if (typeof value === 'number' || value instanceof Date) return 1
    if (typeof value === 'string') return 2
    if (typeof value === 'boolean') return 4
    if (isError(value)) return 16
    if (Array.isArray(value)) return 64
    return 1
  }),

  // -- Dates -----------------------------------------------------------------

  TODAY: def(() => {
    const now = new Date()
    return new Date(now.getFullYear(), now.getMonth(), now.getDate())
  }, { deterministic: false }),
  NOW: def(() => new Date(), { deterministic: false }),
  DATE: def(([year, month, day]) =>
    new Date(Math.trunc(toNumber(year)), Math.trunc(toNumber(month)) - 1, Math.trunc(toNumber(day)))),
  TIME: def(([hour, minute, second]) =>
    ((toNumber(hour) * 3600) + (toNumber(minute) * 60) + toNumber(second)) / 86400),
  DATEVALUE: unary((value) => dateToSerial(toDate(value))),
  YEAR: unary((value) => toDate(value).getFullYear()),
  MONTH: unary((value) => toDate(value).getMonth() + 1),
  DAY: unary((value) => toDate(value).getDate()),
  HOUR: unary((value) => toDate(value).getHours()),
  MINUTE: unary((value) => toDate(value).getMinutes()),
  SECOND: unary((value) => toDate(value).getSeconds()),
  WEEKDAY: def(([value, type]) => {
    const day = toDate(value).getDay()
    const mode = type === undefined ? 1 : Math.trunc(toNumber(type))
    if (mode === 1) return day + 1
    if (mode === 2) return day === 0 ? 7 : day
    if (mode === 3) return day === 0 ? 6 : day - 1
    return day + 1
  }),
  DAYS: def(([end, start]) => Math.round(dateToSerial(toDate(end)) - dateToSerial(toDate(start)))),
  EDATE: def(([value, months]) => {
    const date = toDate(value)
    const shifted = new Date(date.getFullYear(), date.getMonth() + Math.trunc(toNumber(months)), 1)
    const lastDay = new Date(shifted.getFullYear(), shifted.getMonth() + 1, 0).getDate()
    shifted.setDate(Math.min(date.getDate(), lastDay))
    return shifted
  }),
  EOMONTH: def(([value, months]) => {
    const date = toDate(value)
    return new Date(date.getFullYear(), date.getMonth() + Math.trunc(toNumber(months)) + 1, 0)
  }),

  // -- Financial -------------------------------------------------------------

  PMT: def(([rate, periods, present, future, type]) => {
    const r = toNumber(rate)
    const n = toNumber(periods)
    const pv = toNumber(present)
    const fv = future === undefined ? 0 : toNumber(future)
    const due = type === undefined ? 0 : toNumber(type)
    if (n === 0) return NUM
    if (r === 0) return -(pv + fv) / n
    const growth = Math.pow(1 + r, n)
    return finite(-(pv * growth + fv) * r / ((growth - 1) * (1 + r * due)))
  }),
  FV: def(([rate, periods, payment, present, type]) => {
    const r = toNumber(rate)
    const n = toNumber(periods)
    const pmt = toNumber(payment)
    const pv = present === undefined ? 0 : toNumber(present)
    const due = type === undefined ? 0 : toNumber(type)
    if (r === 0) return -(pv + pmt * n)
    const growth = Math.pow(1 + r, n)
    return finite(-(pv * growth + pmt * (1 + r * due) * (growth - 1) / r))
  }),
  PV: def(([rate, periods, payment, future, type]) => {
    const r = toNumber(rate)
    const n = toNumber(periods)
    const pmt = toNumber(payment)
    const fv = future === undefined ? 0 : toNumber(future)
    const due = type === undefined ? 0 : toNumber(type)
    if (r === 0) return -(fv + pmt * n)
    const growth = Math.pow(1 + r, n)
    return finite(-(fv + pmt * (1 + r * due) * (growth - 1) / r) / growth)
  }),
  NPV: def(([rate, ...cashflows]) => {
    const r = toNumber(rate)
    const flows = flatten(cashflows)
    let total = 0
    for (let i = 0; i < flows.length; i++) {
      total += toNumber(flows[i]) / Math.pow(1 + r, i + 1)
    }
    return finite(total)
  })
}

function variance (numbers, sample) {
  const divisor = sample ? numbers.length - 1 : numbers.length
  if (divisor <= 0) return DIV0
  let sum = 0
  for (let i = 0; i < numbers.length; i++) sum += numbers[i]
  const centre = sum / numbers.length
  let squares = 0
  for (let i = 0; i < numbers.length; i++) {
    const deviation = numbers[i] - centre
    squares += deviation * deviation
  }
  return squares / divisor
}

/** Join every (possibly nested) argument as text, for CONCAT/CONCATENATE. */
function concatenate (args) {
  const values = flatten(args)
  let result = ''
  for (let i = 0; i < values.length; i++) result += toText(values[i])
  return result
}

/** Is every row of a grid a single cell wide? */
function isSingleColumn (grid) {
  for (let i = 0; i < grid.length; i++) {
    if (grid[i].length !== 1) return false
  }
  return true
}

/** The leading cell of every row, as a flat vector. */
function firstColumn (grid) {
  const column = new Array(grid.length)
  for (let i = 0; i < grid.length; i++) column[i] = grid[i][0]
  return column
}

function pick (row, index) {
  if (index === 0) return row
  const value = row[index - 1]
  return value === undefined ? REF : value
}

/** Shared row/column scan for VLOOKUP and HLOOKUP. */
function findRow (vector, lookup, approximate) {
  if (!approximate) {
    const matches = matcher(lookup)
    for (let i = 0; i < vector.length; i++) {
      if (matches(vector[i])) return i
    }
    return -1
  }
  let best = -1
  for (let i = 0; i < vector.length; i++) {
    if (isBlank(vector[i])) continue
    if (compareValues(vector[i], lookup) <= 0) best = i
    else break
  }
  return best
}


// ---------------------------------------------------------------------------
// Extended library
//
// Chosen by diffing against HyperFormula's 418 functions and discarding the
// specialist tail (complex numbers, statistical distributions, base conversion,
// bitwise, database and matrix functions). Every entry here is cross-checked
// against HyperFormula by `compare/agreement.js`.
// ---------------------------------------------------------------------------

/** An aggregate taking a range plus one scalar parameter, as PERCENTILE does. */
function aggregate2 (handler) {
  return def(([values, parameter]) => handler(numericOperands([values]), toNumber(parameter)))
}

/** Row indices where every (range, criteria) pair matches. */
function matchingIndices (pairs) {
  const vectors = new Array(pairs.length)
  const predicates = new Array(pairs.length)
  for (let i = 0; i < pairs.length; i++) {
    vectors[i] = flatten([pairs[i][0]])
    predicates[i] = criteriaPredicate(pairs[i][1])
  }
  const length = vectors.length ? vectors[0].length : 0
  const indices = []
  for (let i = 0; i < length; i++) {
    let matched = true
    for (let p = 0; p < predicates.length; p++) {
      if (!predicates[p](vectors[p][i])) { matched = false; break }
    }
    if (matched) indices.push(i)
  }
  return indices
}

/** The numeric entries of a (possibly nested) argument, for cashflow series. */
function numericValues (value) {
  const values = flatten([value])
  const numbers = []
  for (let i = 0; i < values.length; i++) {
    if (typeof values[i] === 'number') numbers.push(values[i])
  }
  return numbers
}

/** Split `range1, criteria1, range2, criteria2, ...` into pairs. */
function criteriaPairs (args) {
  const pairs = []
  for (let i = 0; i + 1 < args.length; i += 2) pairs.push([args[i], args[i + 1]])
  return pairs
}

/** The numbers a *IFS function should consider, at the matching indices. */
function selectedNumbers (target, indices) {
  const values = flatten([target])
  const numbers = []
  for (const index of indices) {
    if (typeof values[index] === 'number') numbers.push(values[index])
  }
  return numbers
}

function sorted (numbers) {
  return [...numbers].sort((a, b) => a - b)
}

function mean (numbers) {
  let total = 0
  for (let i = 0; i < numbers.length; i++) total += numbers[i]
  return total / numbers.length
}

/** Excel's inclusive percentile: linear interpolation across the whole range. */
function percentileInc (numbers, k) {
  if (!numbers.length || k < 0 || k > 1) return NUM
  const list = sorted(numbers)
  const position = k * (list.length - 1)
  const lower = Math.floor(position)
  const fraction = position - lower
  if (lower + 1 >= list.length) return list[list.length - 1]
  return list[lower] + fraction * (list[lower + 1] - list[lower])
}

/** The exclusive variant, which refuses the extreme tails. */
function percentileExc (numbers, k) {
  const list = sorted(numbers)
  const n = list.length
  if (!n) return NUM
  const position = k * (n + 1)
  if (position < 1 || position > n) return NUM
  const lower = Math.floor(position)
  const fraction = position - lower
  if (lower >= n) return list[n - 1]
  return list[lower - 1] + fraction * (list[lower] - list[lower - 1])
}

/** Values as Excel's "A" variants see them: text is 0, logicals are 0/1. */
function operandsA (args) {
  const values = []
  for (const value of flatten(args)) {
    if (isError(value)) throw value
    if (typeof value === 'number') values.push(value)
    else if (typeof value === 'boolean') values.push(value ? 1 : 0)
    else if (value instanceof Date) values.push(dateToSerial(value))
    else if (typeof value === 'string') values.push(0)
  }
  return values
}

/** Paired numeric vectors for the correlation family. */
function pairedNumbers (a, b) {
  const left = flatten([a])
  const right = flatten([b])
  const xs = []
  const ys = []
  for (let i = 0; i < Math.min(left.length, right.length); i++) {
    if (typeof left[i] === 'number' && typeof right[i] === 'number') {
      xs.push(left[i])
      ys.push(right[i])
    }
  }
  return [xs, ys]
}

function pearson (xs, ys) {
  if (xs.length < 2) return DIV0
  const mx = mean(xs)
  const my = mean(ys)
  let sxy = 0
  let sxx = 0
  let syy = 0
  for (let i = 0; i < xs.length; i++) {
    const dx = xs[i] - mx
    const dy = ys[i] - my
    sxy += dx * dy
    sxx += dx * dx
    syy += dy * dy
  }
  if (sxx === 0 || syy === 0) return DIV0
  return sxy / Math.sqrt(sxx * syy)
}

// -- Date-count conventions -------------------------------------------------

/** Excel's 30/360 day count, in US (NASD) or European form. */
function days360Between (start, end, european) {
  let d1 = start.getDate()
  let d2 = end.getDate()
  const m1 = start.getMonth() + 1
  const m2 = end.getMonth() + 1
  const y1 = start.getFullYear()
  const y2 = end.getFullYear()

  if (european) {
    if (d1 === 31) d1 = 30
    if (d2 === 31) d2 = 30
  } else {
    const lastOfMonth = (date) => new Date(date.getFullYear(), date.getMonth() + 1, 0).getDate()
    if (d1 === lastOfMonth(start)) d1 = 30
    if (d2 === 31 && d1 === 30) d2 = 30
  }
  return (y2 - y1) * 360 + (m2 - m1) * 30 + (d2 - d1)
}

const MS_PER_DAY_LOCAL = 86400000
const daysBetween = (start, end) =>
  Math.round((end.getTime() - start.getTime()) / MS_PER_DAY_LOCAL)

const isLeapYear = (year) => (year % 4 === 0 && year % 100 !== 0) || year % 400 === 0

/** Whole working days in [start, end], excluding weekends and holidays. */
function countWorkdays (start, end, holidays) {
  const sign = start > end ? -1 : 1
  const from = sign > 0 ? start : end
  const to = sign > 0 ? end : start
  const excluded = new Set()
  for (let i = 0; i < holidays.length; i++) excluded.add(Math.floor(dateToSerial(holidays[i])))

  let count = 0
  const cursor = new Date(from.getFullYear(), from.getMonth(), from.getDate())
  const limit = new Date(to.getFullYear(), to.getMonth(), to.getDate())
  while (cursor <= limit) {
    const day = cursor.getDay()
    if (day !== 0 && day !== 6 && !excluded.has(Math.floor(dateToSerial(cursor)))) count++
    cursor.setDate(cursor.getDate() + 1)
  }
  return count * sign
}

/** Non-blank entries of an argument, as Dates. Used for holidays and cashflow dates. */
function dateList (value) {
  if (value === undefined || value === null) return []
  const values = flatten([value])
  const dates = []
  for (let i = 0; i < values.length; i++) {
    if (!isBlank(values[i])) dates.push(toDate(values[i]))
  }
  return dates
}

// -- Financial solvers ------------------------------------------------------

/** Future value of an annuity, the identity the whole family is built on. */
function annuityFV (rate, periods, payment, present, type) {
  if (rate === 0) return -(present + payment * periods)
  const growth = Math.pow(1 + rate, periods)
  return -(present * growth + payment * (1 + rate * type) * (growth - 1) / rate)
}

/** Bisection on a sign change, used where Newton is fragile. */
function solve (fn, low, high, tolerance = 1e-13, iterations = 300) {
  let flow = fn(low)
  let fhigh = fn(high)
  if (!Number.isFinite(flow) || !Number.isFinite(fhigh)) return NUM

  // Widen until the bracket straddles a root.
  let attempts = 0
  while (flow * fhigh > 0 && attempts++ < 60) {
    low -= 0.5
    high += 0.5
    flow = fn(low)
    fhigh = fn(high)
    if (!Number.isFinite(flow) || !Number.isFinite(fhigh)) return NUM
  }
  if (flow * fhigh > 0) return NUM

  let mid = low
  for (let i = 0; i < iterations; i++) {
    mid = (low + high) / 2
    const value = fn(mid)
    if (Math.abs(value) < tolerance || (high - low) / 2 < tolerance) return mid
    if (value * flow < 0) { high = mid } else { low = mid; flow = value }
  }
  return mid
}

const ROMAN_NUMERALS = [
  [1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'],
  [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']
]
const ROMAN_VALUES = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 }

Object.assign(functions, {
  // -- Multi-criteria aggregation -------------------------------------------

  COUNTIFS: def((args) => matchingIndices(criteriaPairs(args)).length),
  SUMIFS: def(([sumRange, ...rest]) => {
    const numbers = selectedNumbers(sumRange, matchingIndices(criteriaPairs(rest)))
    let total = 0
    for (let i = 0; i < numbers.length; i++) total += numbers[i]
    return finite(total)
  }),
  AVERAGEIFS: def(([averageRange, ...rest]) => {
    const numbers = selectedNumbers(averageRange, matchingIndices(criteriaPairs(rest)))
    if (!numbers.length) return DIV0
    return mean(numbers)
  }),
  MAXIFS: def(([maxRange, ...rest]) => {
    const numbers = selectedNumbers(maxRange, matchingIndices(criteriaPairs(rest)))
    return numbers.length ? Math.max(...numbers) : 0
  }),
  MINIFS: def(([minRange, ...rest]) => {
    const numbers = selectedNumbers(minRange, matchingIndices(criteriaPairs(rest)))
    return numbers.length ? Math.min(...numbers) : 0
  }),
  COUNTUNIQUE: def((args) => {
    const seen = new Set()
    for (const value of flatten(args)) {
      if (isBlank(value)) continue
      seen.add(typeof value === 'string' ? value.toUpperCase() : value)
    }
    return seen.size
  }),

  /**
   * SUBTOTAL dispatches to another aggregate by number. 101-111 mean "ignore
   * hidden rows", which has no meaning without a UI, so they behave as 1-11.
   */
  SUBTOTAL: def(([selector, ...rest]) => {
    const which = Math.trunc(toNumber(selector)) % 100
    const names = {
      1: 'AVERAGE', 2: 'COUNT', 3: 'COUNTA', 4: 'MAX', 5: 'MIN',
      6: 'PRODUCT', 7: 'STDEV', 8: 'STDEVP', 9: 'SUM', 10: 'VAR', 11: 'VARP'
    }
    const target = names[which]
    if (!target) return VALUE
    return functions[target].method(rest)
  }),

  // -- Statistics -----------------------------------------------------------

  PERCENTILE: aggregate2((numbers, k) => percentileInc(numbers, k)),
  'PERCENTILE.INC': aggregate2((numbers, k) => percentileInc(numbers, k)),
  'PERCENTILE.EXC': aggregate2((numbers, k) => percentileExc(numbers, k)),
  QUARTILE: aggregate2((numbers, q) => percentileInc(numbers, q / 4)),
  'QUARTILE.INC': aggregate2((numbers, q) => percentileInc(numbers, q / 4)),
  'QUARTILE.EXC': aggregate2((numbers, q) => percentileExc(numbers, q / 4)),

  GEOMEAN: aggregate((numbers) => {
    if (!numbers.length) return NUM
    let total = 0
    for (const n of numbers) {
      if (n <= 0) return NUM
      total += Math.log(n)
    }
    return Math.exp(total / numbers.length)
  }),
  HARMEAN: aggregate((numbers) => {
    if (!numbers.length) return NUM
    let total = 0
    for (const n of numbers) {
      if (n <= 0) return NUM
      total += 1 / n
    }
    return numbers.length / total
  }),
  AVEDEV: aggregate((numbers) => {
    if (!numbers.length) return NUM
    const centre = mean(numbers)
    let total = 0
    for (let i = 0; i < numbers.length; i++) total += Math.abs(numbers[i] - centre)
    return total / numbers.length
  }),
  DEVSQ: aggregate((numbers) => {
    if (!numbers.length) return NUM
    const centre = mean(numbers)
    let total = 0
    for (let i = 0; i < numbers.length; i++) {
      const deviation = numbers[i] - centre
      total += deviation * deviation
    }
    return total
  }),
  SUMSQ: aggregate((numbers) => {
    let total = 0
    for (let i = 0; i < numbers.length; i++) total += numbers[i] * numbers[i]
    return finite(total)
  }),

  CORREL: def(([a, b]) => pearson(...pairedNumbers(a, b))),
  RSQ: def(([a, b]) => {
    const r = pearson(...pairedNumbers(a, b))
    return isError(r) ? r : r * r
  }),
  SLOPE: def(([ys, xs]) => {
    const [y, x] = pairedNumbers(ys, xs)
    if (x.length < 2) return DIV0
    const mx = mean(x)
    const my = mean(y)
    let sxy = 0
    let sxx = 0
    for (let i = 0; i < x.length; i++) {
      sxy += (x[i] - mx) * (y[i] - my)
      sxx += (x[i] - mx) ** 2
    }
    if (sxx === 0) return DIV0
    return sxy / sxx
  }),
  INTERCEPT: def(([ys, xs]) => {
    const slope = functions.SLOPE.method([ys, xs])
    if (isError(slope)) return slope
    const [y, x] = pairedNumbers(ys, xs)
    return mean(y) - slope * mean(x)
  }),
  STANDARDIZE: def(([value, centre, deviation]) => {
    const sd = toNumber(deviation)
    if (sd <= 0) return NUM
    return (toNumber(value) - toNumber(centre)) / sd
  }),
  FISHER: unary((n) => {
    const x = toNumber(n)
    if (x <= -1 || x >= 1) return NUM
    return Math.atanh(x)
  }),
  FISHERINV: unary((n) => Math.tanh(toNumber(n))),

  SUMX2MY2: def(([a, b]) => {
    const [x, y] = pairedNumbers(a, b)
    let total = 0
    for (let i = 0; i < x.length; i++) total += x[i] * x[i] - y[i] * y[i]
    return finite(total)
  }),
  SUMX2PY2: def(([a, b]) => {
    const [x, y] = pairedNumbers(a, b)
    let total = 0
    for (let i = 0; i < x.length; i++) total += x[i] * x[i] + y[i] * y[i]
    return finite(total)
  }),
  SUMXMY2: def(([a, b]) => {
    const [x, y] = pairedNumbers(a, b)
    let total = 0
    for (let i = 0; i < x.length; i++) {
      const difference = x[i] - y[i]
      total += difference * difference
    }
    return finite(total)
  }),

  // The "A" variants count text as zero rather than skipping it.
  AVERAGEA: def((args) => {
    const values = operandsA(args)
    return values.length ? mean(values) : DIV0
  }),
  MAXA: def((args) => {
    const values = operandsA(args)
    return values.length ? Math.max(...values) : 0
  }),
  MINA: def((args) => {
    const values = operandsA(args)
    return values.length ? Math.min(...values) : 0
  }),
  VARA: def((args) => variance(operandsA(args), true)),
  VARPA: def((args) => variance(operandsA(args), false)),
  STDEVA: def((args) => {
    const result = variance(operandsA(args), true)
    return isError(result) ? result : Math.sqrt(result)
  }),
  STDEVPA: def((args) => {
    const result = variance(operandsA(args), false)
    return isError(result) ? result : Math.sqrt(result)
  }),

  // -- Maths ----------------------------------------------------------------

  FACT: unary((n) => {
    const value = Math.trunc(toNumber(n))
    if (value < 0) return NUM
    let total = 1
    for (let i = 2; i <= value; i++) total *= i
    return finite(total)
  }),
  FACTDOUBLE: unary((n) => {
    const value = Math.trunc(toNumber(n))
    if (value < -1) return NUM
    let total = 1
    for (let i = value; i > 1; i -= 2) total *= i
    return finite(total)
  }),
  COMBIN: def(([n, k]) => {
    const total = Math.trunc(toNumber(n))
    const chosen = Math.trunc(toNumber(k))
    if (total < 0 || chosen < 0 || chosen > total) return NUM
    let result = 1
    for (let i = 1; i <= chosen; i++) result = (result * (total - chosen + i)) / i
    return finite(Math.round(result))
  }),
  COMBINA: def(([n, k]) => {
    const total = Math.trunc(toNumber(n))
    const chosen = Math.trunc(toNumber(k))
    if (total < 0 || chosen < 0) return NUM
    if (total === 0 && chosen === 0) return 1
    return functions.COMBIN.method([total + chosen - 1, chosen])
  }),
  MULTINOMIAL: aggregate((numbers) => {
    let total = 0
    let result = 1
    for (const n of numbers) {
      if (n < 0) return NUM
      const value = Math.trunc(n)
      for (let i = 1; i <= value; i++) {
        total++
        result = (result * total) / i
      }
    }
    return finite(Math.round(result))
  }),
  QUOTIENT: def(([a, b]) => {
    const divisor = toNumber(b)
    if (divisor === 0) return DIV0
    return Math.trunc(toNumber(a) / divisor)
  }),
  SQRTPI: unary((n) => {
    const value = toNumber(n)
    if (value < 0) return NUM
    return Math.sqrt(value * Math.PI)
  }),
  DELTA: def(([a, b]) => (toNumber(a) === toNumber(b === undefined ? 0 : b) ? 1 : 0)),

  'CEILING.MATH': def(([n, significance, mode]) => {
    const value = toNumber(n)
    const step = significance === undefined ? 1 : Math.abs(toNumber(significance))
    if (step === 0) return 0
    // A non-zero mode rounds negatives away from zero instead of towards it.
    const away = mode !== undefined && toNumber(mode) !== 0
    if (value < 0 && away) return -Math.ceil(Math.abs(value) / step) * step
    return Math.ceil(value / step) * step
  }),
  'FLOOR.MATH': def(([n, significance, mode]) => {
    const value = toNumber(n)
    const step = significance === undefined ? 1 : Math.abs(toNumber(significance))
    if (step === 0) return 0
    const away = mode !== undefined && toNumber(mode) !== 0
    if (value < 0 && away) return -Math.floor(Math.abs(value) / step) * step
    return Math.floor(value / step) * step
  }),
  'CEILING.PRECISE': def(([n, significance]) => {
    const step = significance === undefined ? 1 : Math.abs(toNumber(significance))
    if (step === 0) return 0
    return Math.ceil(toNumber(n) / step) * step
  }),
  'FLOOR.PRECISE': def(([n, significance]) => {
    const step = significance === undefined ? 1 : Math.abs(toNumber(significance))
    if (step === 0) return 0
    return Math.floor(toNumber(n) / step) * step
  }),
  'ISO.CEILING': def(([n, significance]) => {
    const step = significance === undefined ? 1 : Math.abs(toNumber(significance))
    if (step === 0) return 0
    return Math.ceil(toNumber(n) / step) * step
  }),

  BASE: def(([n, radix, minLength]) => {
    const value = Math.trunc(toNumber(n))
    const base = Math.trunc(toNumber(radix))
    if (value < 0 || base < 2 || base > 36) return NUM
    const text = value.toString(base).toUpperCase()
    const pad = minLength === undefined ? 0 : Math.trunc(toNumber(minLength))
    return text.padStart(pad, '0')
  }),
  DECIMAL: def(([text, radix]) => {
    const base = Math.trunc(toNumber(radix))
    if (base < 2 || base > 36) return NUM
    const source = toText(text).trim()
    if (!source) return NUM
    const parsed = parseInt(source, base)
    if (Number.isNaN(parsed)) return NUM
    return parsed
  }),
  ROMAN: def(([n]) => {
    let value = Math.trunc(toNumber(n))
    if (value < 0 || value > 3999) return VALUE
    let result = ''
    for (const [amount, numeral] of ROMAN_NUMERALS) {
      while (value >= amount) { result += numeral; value -= amount }
    }
    return result
  }),
  ARABIC: unary((text) => {
    const source = toText(text).toUpperCase().trim()
    if (!source) return 0
    const negative = source.startsWith('-')
    const body = negative ? source.slice(1) : source
    if (!/^[IVXLCDM]*$/.test(body)) return VALUE
    let total = 0
    for (let i = 0; i < body.length; i++) {
      const current = ROMAN_VALUES[body[i]]
      const next = ROMAN_VALUES[body[i + 1]]
      total += next > current ? -current : current
    }
    return negative ? -total : total
  }),

  // -- Trigonometry ---------------------------------------------------------

  ASINH: unary((n) => Math.asinh(toNumber(n))),
  ACOSH: unary((n) => {
    const value = toNumber(n)
    if (value < 1) return NUM
    return Math.acosh(value)
  }),
  ATANH: unary((n) => {
    const value = toNumber(n)
    if (value <= -1 || value >= 1) return NUM
    return Math.atanh(value)
  }),
  COT: unary((n) => {
    const value = toNumber(n)
    if (value === 0) return DIV0
    return 1 / Math.tan(value)
  }),
  COTH: unary((n) => {
    const value = toNumber(n)
    if (value === 0) return DIV0
    return 1 / Math.tanh(value)
  }),
  ACOT: unary((n) => Math.PI / 2 - Math.atan(toNumber(n))),
  ACOTH: unary((n) => {
    const value = toNumber(n)
    if (Math.abs(value) <= 1) return NUM
    return Math.atanh(1 / value)
  }),
  SEC: unary((n) => 1 / Math.cos(toNumber(n))),
  SECH: unary((n) => 1 / Math.cosh(toNumber(n))),
  CSC: unary((n) => {
    const value = toNumber(n)
    if (value === 0) return DIV0
    return 1 / Math.sin(value)
  }),
  CSCH: unary((n) => {
    const value = toNumber(n)
    if (value === 0) return DIV0
    return 1 / Math.sinh(value)
  }),

  // -- Text -----------------------------------------------------------------

  SPLIT: def(([text, delimiter]) => {
    const separator = delimiter === undefined ? ' ' : toText(delimiter)
    return toText(text).split(separator)
  }),
  UNICHAR: unary((code) => {
    const value = Math.trunc(toNumber(code))
    if (value < 1 || value > 0x10FFFF) return VALUE
    return String.fromCodePoint(value)
  }),
  UNICODE: unary((text) => {
    const source = toText(text)
    if (!source.length) return VALUE
    return source.codePointAt(0)
  }),

  // -- Dates ----------------------------------------------------------------

  TIMEVALUE: unary((text) => {
    const match = /(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(am|pm)?/i.exec(toText(text))
    if (!match) return VALUE
    let hours = Number(match[1])
    const minutes = Number(match[2])
    const seconds = match[3] ? Number(match[3]) : 0
    const meridiem = match[4] && match[4].toLowerCase()
    if (meridiem === 'pm' && hours < 12) hours += 12
    if (meridiem === 'am' && hours === 12) hours = 0
    return (hours * 3600 + minutes * 60 + seconds) / 86400
  }),
  DAYS360: def(([start, end, method]) =>
    days360Between(toDate(start), toDate(end), method !== undefined && toBoolean(method))),
  DATEDIF: def(([start, end, unit]) => {
    const from = toDate(start)
    const to = toDate(end)
    if (from > to) return NUM
    const code = toText(unit).toUpperCase()

    let years = to.getFullYear() - from.getFullYear()
    let months = to.getMonth() - from.getMonth()
    let days = to.getDate() - from.getDate()
    if (days < 0) {
      months--
      days += new Date(to.getFullYear(), to.getMonth(), 0).getDate()
    }
    if (months < 0) { months += 12; years-- }

    switch (code) {
      case 'Y': return years
      case 'M': return years * 12 + months
      case 'D': return daysBetween(from, to)
      case 'MD': return days
      case 'YM': return months
      case 'YD': {
        // Days ignoring years: advance the start date to the same year as the end.
        const anniversary = new Date(from.getFullYear() + years, from.getMonth(), from.getDate())
        return daysBetween(anniversary, to)
      }
      default: return NUM
    }
  }),
  YEARFRAC: def(([start, end, basis]) => {
    let from = toDate(start)
    let to = toDate(end)
    if (from > to) [from, to] = [to, from]
    const which = basis === undefined ? 0 : Math.trunc(toNumber(basis))

    switch (which) {
      case 0: return days360Between(from, to, false) / 360
      case 1: {
        // Actual/actual: denominator is the average year length across the span.
        const years = to.getFullYear() - from.getFullYear() + 1
        let days = 0
        for (let year = from.getFullYear(); year <= to.getFullYear(); year++) {
          days += isLeapYear(year) ? 366 : 365
        }
        return daysBetween(from, to) / (days / years)
      }
      case 2: return daysBetween(from, to) / 360
      case 3: return daysBetween(from, to) / 365
      case 4: return days360Between(from, to, true) / 360
      default: return NUM
    }
  }),
  WEEKNUM: def(([value, type]) => {
    const date = toDate(value)
    const mode = type === undefined ? 1 : Math.trunc(toNumber(type))
    if (mode === 21) return functions.ISOWEEKNUM.method(date)
    // Which weekday starts the week, as a JS day number.
    const starts = { 1: 0, 2: 1, 11: 1, 12: 2, 13: 3, 14: 4, 15: 5, 16: 6, 17: 0 }
    const startDay = starts[mode]
    if (startDay === undefined) return NUM
    const firstOfYear = new Date(date.getFullYear(), 0, 1)
    const offset = (firstOfYear.getDay() - startDay + 7) % 7
    const dayOfYear = daysBetween(firstOfYear, date) + 1
    return Math.floor((dayOfYear + offset - 1) / 7) + 1
  }),
  ISOWEEKNUM: unary((value) => {
    const date = toDate(value)
    const target = new Date(date.getFullYear(), date.getMonth(), date.getDate())
    // Shift to the Thursday of this ISO week; its year owns the week.
    const dayNumber = (target.getDay() + 6) % 7
    target.setDate(target.getDate() - dayNumber + 3)
    const firstThursday = new Date(target.getFullYear(), 0, 4)
    const firstDayNumber = (firstThursday.getDay() + 6) % 7
    firstThursday.setDate(firstThursday.getDate() - firstDayNumber + 3)
    return 1 + Math.round(daysBetween(firstThursday, target) / 7)
  }),
  NETWORKDAYS: def(([start, end, holidays]) =>
    countWorkdays(toDate(start), toDate(end), dateList(holidays))),
  WORKDAY: def(([start, days, holidays]) => {
    const dates = dateList(holidays)
    const excluded = new Set()
    for (let i = 0; i < dates.length; i++) excluded.add(Math.floor(dateToSerial(dates[i])))
    let remaining = Math.trunc(toNumber(days))
    const step = remaining < 0 ? -1 : 1
    const date = toDate(start)
    const cursor = new Date(date.getFullYear(), date.getMonth(), date.getDate())
    while (remaining !== 0) {
      cursor.setDate(cursor.getDate() + step)
      const day = cursor.getDay()
      if (day !== 0 && day !== 6 && !excluded.has(Math.floor(dateToSerial(cursor)))) {
        remaining -= step
      }
    }
    return cursor
  }),

  // -- Financial ------------------------------------------------------------

  NPER: def(([rate, payment, present, future, type]) => {
    const r = toNumber(rate)
    const pmt = toNumber(payment)
    const pv = toNumber(present)
    const fv = future === undefined ? 0 : toNumber(future)
    const due = type === undefined ? 0 : toNumber(type)
    if (r === 0) {
      if (pmt === 0) return NUM
      return -(pv + fv) / pmt
    }
    const adjusted = pmt * (1 + r * due)
    const numerator = adjusted - fv * r
    const denominator = pv * r + adjusted
    if (numerator / denominator <= 0) return NUM
    return Math.log(numerator / denominator) / Math.log(1 + r)
  }),
  RATE: def(([periods, payment, present, future, type, guess]) => {
    const n = toNumber(periods)
    const pmt = toNumber(payment)
    const pv = toNumber(present)
    const fv = future === undefined ? 0 : toNumber(future)
    const due = type === undefined ? 0 : toNumber(type)
    void guess
    const result = solve((rate) => annuityFV(rate, n, pmt, pv, due) - fv, -0.99, 1)
    return isError(result) ? result : result
  }),
  IPMT: def(([rate, period, periods, present, future, type]) => {
    const r = toNumber(rate)
    const per = Math.trunc(toNumber(period))
    const n = toNumber(periods)
    if (per < 1 || per > n) return NUM
    const payment = functions.PMT.method([rate, periods, present, future, type])
    if (isError(payment)) return payment
    const due = type === undefined ? 0 : toNumber(type)
    // Balance owing at the start of this period, which is negative for a loan —
    // so the interest charged on it is negative too, matching Excel's convention
    // that money leaving your pocket is a negative number.
    const balance = annuityFV(r, per - 1, payment, toNumber(present), due)
    if (per === 1 && due === 1) return 0
    return finite(balance * r * (due === 1 ? 1 / (1 + r) : 1))
  }),
  PPMT: def(([rate, period, periods, present, future, type]) => {
    const payment = functions.PMT.method([rate, periods, present, future, type])
    if (isError(payment)) return payment
    const interest = functions.IPMT.method([rate, period, periods, present, future, type])
    if (isError(interest)) return interest
    return finite(payment - interest)
  }),
  ISPMT: def(([rate, period, periods, present]) => {
    const r = toNumber(rate)
    const per = toNumber(period)
    const n = toNumber(periods)
    const pv = toNumber(present)
    return finite(pv * r * (per / n - 1))
  }),
  CUMIPMT: def(([rate, periods, present, start, end, type]) => {
    const from = Math.trunc(toNumber(start))
    const to = Math.trunc(toNumber(end))
    if (from < 1 || to < from) return NUM
    let total = 0
    for (let period = from; period <= to; period++) {
      const interest = functions.IPMT.method([rate, period, periods, present, 0, type])
      if (isError(interest)) return interest
      total += interest
    }
    return finite(total)
  }),
  CUMPRINC: def(([rate, periods, present, start, end, type]) => {
    const from = Math.trunc(toNumber(start))
    const to = Math.trunc(toNumber(end))
    if (from < 1 || to < from) return NUM
    let total = 0
    for (let period = from; period <= to; period++) {
      const principal = functions.PPMT.method([rate, period, periods, present, 0, type])
      if (isError(principal)) return principal
      total += principal
    }
    return finite(total)
  }),
  IRR: def(([values, guess]) => {
    const flows = numericValues(values)
    if (flows.length < 2) return NUM
    void guess
    const npv = (rate) => {
      let total = 0
      for (let i = 0; i < flows.length; i++) total += flows[i] / Math.pow(1 + rate, i)
      return total
    }
    return solve(npv, -0.99, 1)
  }),
  MIRR: def(([values, financeRate, reinvestRate]) => {
    const flows = numericValues(values)
    const finance = toNumber(financeRate)
    const reinvest = toNumber(reinvestRate)
    const n = flows.length
    if (n < 2) return DIV0
    let negatives = 0
    let positives = 0
    for (let i = 0; i < n; i++) {
      if (flows[i] < 0) negatives += flows[i] / Math.pow(1 + finance, i)
      else positives += flows[i] * Math.pow(1 + reinvest, n - 1 - i)
    }
    if (negatives === 0) return DIV0
    return finite(Math.pow(-positives / negatives, 1 / (n - 1)) - 1)
  }),
  XNPV: def(([rate, values, dates]) => {
    const r = toNumber(rate)
    const flows = numericValues(values)
    const when = dateList(dates)
    if (flows.length !== when.length || !flows.length) return NUM
    const start = when[0]
    let total = 0
    for (let i = 0; i < flows.length; i++) {
      total += flows[i] / Math.pow(1 + r, daysBetween(start, when[i]) / 365)
    }
    return finite(total)
  }),
  FVSCHEDULE: def(([principal, schedule]) => {
    let total = toNumber(principal)
    for (const rate of flatten([schedule])) {
      if (isBlank(rate)) continue
      total *= 1 + toNumber(rate)
    }
    return finite(total)
  }),
  PDURATION: def(([rate, present, future]) => {
    const r = toNumber(rate)
    const pv = toNumber(present)
    const fv = toNumber(future)
    if (r <= 0 || pv <= 0 || fv <= 0) return NUM
    return (Math.log(fv) - Math.log(pv)) / Math.log(1 + r)
  }),
  RRI: def(([periods, present, future]) => {
    const n = toNumber(periods)
    const pv = toNumber(present)
    const fv = toNumber(future)
    if (n <= 0 || pv <= 0) return NUM
    return Math.pow(fv / pv, 1 / n) - 1
  }),
  SLN: def(([cost, salvage, life]) => {
    const years = toNumber(life)
    if (years === 0) return DIV0
    return (toNumber(cost) - toNumber(salvage)) / years
  }),
  SYD: def(([cost, salvage, life, period]) => {
    const years = toNumber(life)
    const per = toNumber(period)
    if (years <= 0) return NUM
    if (per < 1 || per > years) return NUM
    return ((toNumber(cost) - toNumber(salvage)) * (years - per + 1) * 2) / (years * (years + 1))
  }),
  DDB: def(([cost, salvage, life, period, factor]) => {
    const start = toNumber(cost)
    const residual = toNumber(salvage)
    const years = toNumber(life)
    const per = toNumber(period)
    const rate = factor === undefined ? 2 : toNumber(factor)
    if (years <= 0 || per < 1 || per > years) return NUM
    let book = start
    let depreciation = 0
    for (let i = 1; i <= per; i++) {
      depreciation = Math.min((book * rate) / years, book - residual)
      if (depreciation < 0) depreciation = 0
      book -= depreciation
    }
    return finite(depreciation)
  }),
  DB: def(([cost, salvage, life, period, month]) => {
    const start = toNumber(cost)
    const residual = toNumber(salvage)
    const years = toNumber(life)
    const per = toNumber(period)
    const months = month === undefined ? 12 : toNumber(month)
    if (years <= 0 || per < 1) return NUM
    const rate = Number((1 - Math.pow(residual / start, 1 / years)).toFixed(3))
    let book = start
    let depreciation = (start * rate * months) / 12
    if (per === 1) return finite(depreciation)
    book -= depreciation
    for (let i = 2; i <= per; i++) {
      if (i === Math.ceil(years) + 1) {
        depreciation = (book * rate * (12 - months)) / 12
      } else {
        depreciation = book * rate
      }
      book -= depreciation
    }
    return finite(depreciation)
  }),
  EFFECT: def(([nominal, periods]) => {
    const rate = toNumber(nominal)
    const n = Math.trunc(toNumber(periods))
    if (rate <= 0 || n < 1) return NUM
    return Math.pow(1 + rate / n, n) - 1
  }),
  NOMINAL: def(([effective, periods]) => {
    const rate = toNumber(effective)
    const n = Math.trunc(toNumber(periods))
    if (rate <= 0 || n < 1) return NUM
    return (Math.pow(rate + 1, 1 / n) - 1) * n
  }),
  DOLLARDE: def(([fractional, fraction]) => {
    const denominator = Math.trunc(toNumber(fraction))
    if (denominator < 0) return NUM
    if (denominator === 0) return DIV0
    const value = toNumber(fractional)
    const whole = Math.trunc(value)
    return whole + ((value - whole) * Math.pow(10, Math.ceil(Math.log10(denominator)))) / denominator
  }),
  DOLLARFR: def(([decimal, fraction]) => {
    const denominator = Math.trunc(toNumber(fraction))
    if (denominator < 0) return NUM
    if (denominator === 0) return DIV0
    const value = toNumber(decimal)
    const whole = Math.trunc(value)
    return whole + ((value - whole) * denominator) / Math.pow(10, Math.ceil(Math.log10(denominator)))
  })
})

// Modern Excel spellings of functions we already provide.
for (const [alias, target] of Object.entries({
  'VAR.S': 'VAR', 'VAR.P': 'VARP', VARS: 'VAR',
  'STDEV.S': 'STDEV', 'STDEV.P': 'STDEVP', STDEVS: 'STDEV'
})) {
  functions[alias] = functions[target]
}

// Unary negation compiles to `{ '-': [x] }`, so `-` benefits from the same
// treatment. Both stay variadic: the engine only unwraps a single argument, and
// `def` accepts either shape.
functions['-'].optimizeUnary = true
functions['+'].optimizeUnary = true

/** Names the compiler will accept in a formula. */
export const FUNCTION_NAMES = Object.keys(functions).filter((name) => /^[A-Z]/.test(name))

/**
 * Functions that take their argument unwrapped. The grammar reads this to emit
 * `{ ABS: x }` instead of `{ ABS: [x] }`, so the engine never has to undo a
 * wrapper the compiler should not have added in the first place.
 */
export const UNARY_FUNCTION_NAMES = FUNCTION_NAMES.filter(
  (name) => functions[name].optimizeUnary === true
)

export { FormulaError, isError, indexToColumn, parseA1, serialToDate }
