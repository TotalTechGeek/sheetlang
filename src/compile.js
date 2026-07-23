/**
 * Formula text -> JSON Logic.
 *
 * The parser generated from `grammar.peggy` already emits JSON Logic, so this
 * module is only responsible for supplying the known-function set and turning
 * Peggy's SyntaxError into something with a readable message.
 */

import * as parser from './parser.js'
import { functions, UNARY_FUNCTION_NAMES } from './functions.js'
import { FormulaSyntaxError } from './errors.js'

const DEFAULT_FUNCTION_NAMES = new Set(
  Object.keys(functions).filter((name) => /^[A-Z]/.test(name))
)

/**
 * Functions declared `optimizeUnary`, for which the grammar emits the collapsed
 * `{ ABS: x }` form. Derived from the library rather than listed separately, so the
 * two can never drift apart.
 */
const DEFAULT_UNARY_NAMES = new Set(UNARY_FUNCTION_NAMES)

/**
 * Compile a formula into a JSON Logic tree.
 *
 * @param {string} source e.g. `=SUM(A1:A10) * (1 + taxRate)`
 * @param {{ functions?: Iterable<string> | null, unaryFunctions?: Iterable<string> }} [options]
 *   `functions` overrides the set of names accepted; pass `null` to skip the
 *   check entirely (useful when the engine is extended at runtime).
 *   `unaryFunctions` names the functions to emit in the collapsed `{ FN: x }` form;
 *   supply it when a custom engine declares its own `optimizeUnary` methods.
 * @returns {any} JSON Logic
 */
export function compile (source, options = {}) {
  if (typeof source !== 'string') {
    throw new FormulaSyntaxError('A formula must be a string')
  }

  const names = options.functions === null
    ? null
    : options.functions
      ? new Set([...options.functions].map((name) => name.toUpperCase()))
      : DEFAULT_FUNCTION_NAMES

  const unary = options.unaryFunctions
    ? new Set([...options.unaryFunctions].map((name) => name.toUpperCase()))
    : DEFAULT_UNARY_NAMES

  try {
    return parser.parse(source, { functions: names, unaryFunctions: unary })
  } catch (error) {
    if (error && error.name === 'SyntaxError' && error.location) {
      throw new FormulaSyntaxError(describe(source, error), error.location)
    }
    throw error
  }
}

/** Peggy's message plus a caret pointing at the offending column. */
function describe (source, error) {
  const { start } = error.location
  const line = source.split('\n')[start.line - 1] ?? source
  const caret = ' '.repeat(Math.max(0, start.column - 1)) + '^'
  return `${error.message} at column ${start.column}\n  ${line}\n  ${caret}`
}

export { DEFAULT_FUNCTION_NAMES }
