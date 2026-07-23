/**
 * Excel's coercion rules, isolated here so the function library can stay declarative.
 *
 * The rules that matter:
 *  - blank (null/undefined) is 0 in a numeric context, "" in a text context, FALSE in a logical one
 *  - TRUE/FALSE are 1/0 when a number is wanted
 *  - text is coerced to a number only if the whole string parses; otherwise #VALUE!
 *  - Date objects round-trip through Excel serial numbers (days since 1899-12-30)
 */

import { isError, VALUE, NUM } from './errors.js'

/** Excel's epoch: serial 1 is 1900-01-01, and serial 0 is the fictional 1899-12-30. */
const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30)
const MS_PER_DAY = 86400000

export function isBlank (value) {
  return value === null || value === undefined || value === ''
}

export function dateToSerial (date) {
  return (date.getTime() - EXCEL_EPOCH_MS) / MS_PER_DAY
}

export function serialToDate (serial) {
  return new Date(Math.round(serial * MS_PER_DAY) + EXCEL_EPOCH_MS)
}

/**
 * Coerce to a number, the way Excel does when a formula needs one.
 * @throws {FormulaError} #VALUE! when the value is text that is not numeric.
 */
export function toNumber (value) {
  if (isError(value)) throw value
  if (value === null || value === undefined || value === '') return 0
  if (typeof value === 'number') {
    if (Number.isNaN(value)) throw NUM
    return value
  }
  if (typeof value === 'boolean') return value ? 1 : 0
  if (value instanceof Date) return dateToSerial(value)
  if (typeof value === 'string') {
    const trimmed = value.trim()
    if (trimmed === '') return 0
    // Accept a trailing % the way Excel's text-to-number conversion does.
    if (trimmed.endsWith('%')) {
      const pct = Number(trimmed.slice(0, -1).trim())
      if (!Number.isNaN(pct)) return pct / 100
    }
    const n = Number(trimmed)
    if (Number.isNaN(n)) throw VALUE
    return n
  }
  throw VALUE
}

/** Coerce to text. Booleans uppercase, blanks empty, numbers use JS formatting. */
export function toText (value) {
  if (isError(value)) throw value
  if (value === null || value === undefined) return ''
  if (typeof value === 'string') return value
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE'
  if (value instanceof Date) return String(dateToSerial(value))
  if (typeof value === 'number') {
    if (Number.isNaN(value)) throw NUM
    if (!Number.isFinite(value)) throw NUM
    return String(value)
  }
  return String(value)
}

/** Coerce to a boolean. Text is only truthy for the literals TRUE/FALSE. */
export function toBoolean (value) {
  if (isError(value)) throw value
  if (value === null || value === undefined || value === '') return false
  if (typeof value === 'boolean') return value
  if (typeof value === 'number') return value !== 0
  if (value instanceof Date) return true
  if (typeof value === 'string') {
    const upper = value.trim().toUpperCase()
    if (upper === 'TRUE') return true
    if (upper === 'FALSE') return false
    const n = Number(value)
    if (!Number.isNaN(n)) return n !== 0
    throw VALUE
  }
  throw VALUE
}

/** A date-ish value as a JS Date, accepting serials, ISO strings and Dates. */
export function toDate (value) {
  if (isError(value)) throw value
  if (value instanceof Date) return value
  if (typeof value === 'string') {
    const parsed = Date.parse(value)
    if (!Number.isNaN(parsed)) return new Date(parsed)
  }
  const serial = toNumber(value)
  if (serial < 0) throw NUM
  return serialToDate(serial)
}

/**
 * Flatten range/array arguments into a single list, the way SUM(A1:B2, C1) does.
 * Errors inside the list are preserved so callers can decide to propagate.
 */
export function flatten (values, out = []) {
  for (const value of values) {
    if (Array.isArray(value)) flatten(value, out)
    else out.push(value)
  }
  return out
}

/** The first error found anywhere in the (possibly nested) arguments, or undefined. */
export function findError (values) {
  for (const value of values) {
    if (isError(value)) return value
    if (Array.isArray(value)) {
      const nested = findError(value)
      if (nested) return nested
    }
  }
  return undefined
}

/**
 * The numbers Excel's statistical functions actually see: nested values are
 * flattened, and text/blanks/booleans coming *out of a range* are skipped rather
 * than coerced. Direct scalar arguments are coerced, matching SUM(TRUE) === 1.
 */
export function numericOperands (args) {
  const numbers = []
  for (const arg of args) {
    if (Array.isArray(arg)) {
      for (const value of flatten([arg])) {
        if (typeof value === 'number') numbers.push(value)
        else if (value instanceof Date) numbers.push(dateToSerial(value))
      }
    } else if (!isBlank(arg) || arg === false) {
      // A blank scalar contributes nothing, which is why AVERAGE of one is #DIV/0!.
      numbers.push(toNumber(arg))
    }
  }
  return numbers
}

/**
 * Excel's comparison order: numbers < text < FALSE < TRUE, text compared
 * case-insensitively, blanks coerced to match the other operand.
 * @returns {number} negative, zero or positive
 */
export function compareValues (a, b) {
  if (isError(a)) throw a
  if (isError(b)) throw b
  if (a instanceof Date) a = dateToSerial(a)
  if (b instanceof Date) b = dateToSerial(b)

  const aBlank = isBlank(a) && a !== ''
  const bBlank = isBlank(b) && b !== ''
  if (aBlank && bBlank) return 0
  // A blank takes on the type of whatever it is compared against.
  if (aBlank) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0
  if (bBlank) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0

  const rank = (v) => (typeof v === 'boolean' ? 2 : typeof v === 'string' ? 1 : 0)
  const aRank = rank(a)
  const bRank = rank(b)
  if (aRank !== bRank) return aRank - bRank

  if (aRank === 1) {
    const au = String(a).toUpperCase()
    const bu = String(b).toUpperCase()
    return au < bu ? -1 : au > bu ? 1 : 0
  }
  if (aRank === 2) return (a ? 1 : 0) - (b ? 1 : 0)
  const an = toNumber(a)
  const bn = toNumber(b)
  return an < bn ? -1 : an > bn ? 1 : 0
}
