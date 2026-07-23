/**
 * Excel-style error values.
 *
 * Errors are *values*, not exceptions — that is how a spreadsheet behaves, and it
 * keeps the JSON Logic tree free of try/throw plumbing. A function that receives an
 * error argument propagates it; `IFERROR` is the only thing that swallows one.
 */

export class FormulaError extends Error {
  /**
   * @param {string} code The Excel error token, e.g. `#DIV/0!`
   * @param {string} [detail] Human-readable explanation, for debugging.
   */
  constructor (code, detail) {
    super(detail ? `${code}: ${detail}` : code)
    this.name = 'FormulaError'
    this.code = code
    this.detail = detail
  }

  toString () { return this.code }
  toJSON () { return { error: this.code } }
}

/** The canonical error singletons, keyed by their Excel token. */
export const ERROR = {
  '#NULL!': new FormulaError('#NULL!', 'ranges do not intersect'),
  '#DIV/0!': new FormulaError('#DIV/0!', 'division by zero'),
  '#VALUE!': new FormulaError('#VALUE!', 'wrong type of argument'),
  '#REF!': new FormulaError('#REF!', 'invalid reference'),
  '#NAME?': new FormulaError('#NAME?', 'unrecognised name'),
  '#NUM!': new FormulaError('#NUM!', 'invalid numeric value'),
  '#N/A': new FormulaError('#N/A', 'value not available'),
  '#CIRCULAR!': new FormulaError('#CIRCULAR!', 'circular reference')
}

export const NULL_ERROR = ERROR['#NULL!']
export const DIV0 = ERROR['#DIV/0!']
export const VALUE = ERROR['#VALUE!']
export const REF = ERROR['#REF!']
export const NAME = ERROR['#NAME?']
export const NUM = ERROR['#NUM!']
export const NA = ERROR['#N/A']
export const CIRCULAR = ERROR['#CIRCULAR!']

/** Every error token the grammar accepts as a literal. */
export const ERROR_CODES = Object.keys(ERROR)

export function isError (value) {
  return value instanceof FormulaError
}

/**
 * Raised by `compile()` when the source text cannot be turned into logic.
 * Carries Peggy's location info when the failure came from the parser.
 */
export class FormulaSyntaxError extends Error {
  constructor (message, location) {
    super(message)
    this.name = 'FormulaSyntaxError'
    this.location = location
  }
}
