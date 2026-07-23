/**
 * The execution side: a `LogicEngine` preloaded with the Excel function library.
 *
 * Nothing here is special-cased for the compiler — the engine will happily run any
 * JSON Logic you hand it, including hand-written trees that mix Excel functions
 * with the engine's own operations.
 */

import { LogicEngine } from 'json-logic-engine'
import { functions } from './functions.js'

/**
 * @param {{ functions?: Record<string, any>, [key: string]: any }} [options]
 *   `functions` adds to (or overrides) the standard library; every other option is
 *   forwarded to `LogicEngine`.
 */
export function createEngine (options = {}) {
  const { functions: extra = {}, ...engineOptions } = options
  const engine = new LogicEngine(undefined, engineOptions)
  for (const [name, method] of Object.entries({ ...functions, ...extra })) {
    engine.addMethod(name, method)
  }
  return engine
}

/** The engine used when a caller does not supply one. */
export const defaultEngine = createEngine()
