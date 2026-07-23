/**
 * An Excel-like formula language: a Peggy grammar compiles to JSON Logic, and
 * json-logic-engine executes it.
 *
 *   import { evaluate } from 'formulas'
 *   evaluate('=SUM(A1:A3) * (1 + tax)', { A1: 1, A2: 2, A3: 3, tax: 0.1 }) // 6.6
 */

import { compile } from './compile.js'
import { defaultEngine, createEngine } from './engine.js'
import { functions } from './functions.js'

/**
 * Compile once, then run repeatedly. Uses the engine's compiler, so the returned
 * function is considerably faster than `evaluate` in a loop.
 *
 * Pass `{ interpreted: true }` to skip the compiler and interpret on every call.
 * The compiler emits JavaScript via `eval`, which a strict Content-Security-Policy
 * (no `unsafe-eval`) forbids; interpreted mode never generates code, at the cost of
 * some speed. See the CSP note in the README.
 *
 * @param {string} source
 * @param {{ engine?: any, functions?: Iterable<string> | null, interpreted?: boolean }} [options]
 * @returns {(data?: any) => any}
 */
export function build (source, options = {}) {
  const engine = options.engine ?? defaultEngine
  const logic = compile(source, options)
  if (options.interpreted) {
    return (data = {}) => {
      const result = engine.run(logic, data)
      return result === undefined ? null : result
    }
  }
  const run = engine.build(logic)
  return (data = {}) => {
    const result = run(data)
    return result === undefined ? null : result
  }
}

/**
 * Compile and run a formula in one step.
 *
 * @param {string} source
 * @param {any} [data] The variables the formula may reference.
 * @param {{ engine?: any, functions?: Iterable<string> | null }} [options]
 */
export function evaluate (source, data = {}, options = {}) {
  const engine = options.engine ?? defaultEngine
  const result = engine.run(compile(source, options), data)
  return result === undefined ? null : result
}

/** Every function name the compiler accepts. */
export const functionNames = Object.keys(functions)
  .filter((name) => /^[A-Z]/.test(name))
  .sort()

export { compile, createEngine, defaultEngine, functions }
export { Sheet } from './sheet.js'
export {
  FormulaError, FormulaSyntaxError, ERROR, ERROR_CODES, isError
} from './errors.js'
export { dateToSerial, serialToDate } from './coerce.js'
export { expandRange, parseA1, columnToIndex, indexToColumn } from './refs.js'
export { precedentsOf } from './precedents.js'
