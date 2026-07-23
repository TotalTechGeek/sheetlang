/**
 * A minimal spreadsheet on top of the formula compiler.
 *
 * Cells hold either a literal value or a formula string starting with `=`.
 * Evaluation is lazy and memoised: reading a cell evaluates whatever it depends
 * on, through a Proxy that stands in for the data object handed to the engine.
 * Re-entering a cell that is already being evaluated yields `#CIRCULAR!`.
 *
 * Recalculation is incremental. Because the compiled logic is data, each formula's
 * precedents are read straight off its tree at write time (see `precedents.js`) and
 * kept in a reverse index, so a `set` invalidates only the affected subgraph rather
 * than the whole sheet. Pass `{ incremental: false }` to fall back to clearing
 * everything — slower, but the simplest possible behaviour, and the oracle the
 * differential test checks the index against.
 */

import { compile } from './compile.js'
import { defaultEngine } from './engine.js'
import { precedentsOf, rectContains } from './precedents.js'
import { CIRCULAR, isError, FormulaSyntaxError, ERROR } from './errors.js'

export class Sheet {
  /**
   * @param {Record<string, any>} [cells] Initial cells, e.g. `{ A1: 5, B1: '=A1*2' }`
   * @param {{ engine?: any, incremental?: boolean, interpreted?: boolean }} [options]
   */
  constructor (cells = {}, options = {}) {
    this.engine = options.engine ?? defaultEngine
    this.incremental = options.incremental !== false
    // When set, cells are always interpreted and never compiled with the engine's
    // code generator, so the sheet runs under a strict Content-Security-Policy.
    this.interpreted = options.interpreted === true

    /** @type {Map<string, { logic: any, run: Function|null, evaluations: number } | { value: any }>} */
    this.cells = new Map()
    this.cache = new Map()
    this.evaluating = new Set()

    // name -> the cells that read it directly
    /** @type {Map<string, Set<string>>} */
    this.dependents = new Map()
    // Ranges stay as rectangles: expanding `SUM(A1:E200)` would put a thousand
    // entries in the index for one formula.
    /** @type {{ name: string, top: number, left: number, bottom: number, right: number }[]} */
    this.rangeDependents = []
    // Cells whose references could not be resolved statically; always invalidated.
    /** @type {Set<string>} */
    this.dynamic = new Set()

    // Reads against this proxy resolve cells on demand, so a formula referring to
    // `B1` transparently triggers B1's own evaluation.
    this.data = new Proxy({}, {
      get: (_target, key) => (typeof key === 'string' ? this.get(key) : undefined),
      has: (_target, key) => this.cells.has(key),
      ownKeys: () => [...this.cells.keys()],
      getOwnPropertyDescriptor: () => ({ enumerable: true, configurable: true })
    })

    this.setAll(cells)
  }

  /**
   * Store a cell. A string beginning with `=` is compiled immediately, so syntax
   * errors surface at write time rather than at read time.
   */
  set (name, value) {
    this.unindex(name)

    if (typeof value === 'string' && value.startsWith('=')) {
      // Compiled now, so syntax errors surface at write time. Building is deferred
      // to `get` — see there for why.
      const logic = compile(value)
      const precedents = precedentsOf(logic)
      this.cells.set(name, { logic, run: null, evaluations: 0, precedents })
      this.index(name, precedents)
    } else {
      this.cells.set(name, { value })
    }

    this.invalidate(name)
    return this
  }

  setAll (cells) {
    for (const [name, value] of Object.entries(cells)) this.set(name, value)
    return this
  }

  delete (name) {
    this.unindex(name)
    this.cells.delete(name)
    this.invalidate(name)
    return this
  }

  /** The evaluated value of a cell; `null` for one that was never set. */
  get (name) {
    if (this.cache.has(name)) return this.cache.get(name)

    const cell = this.cells.get(name)
    if (!cell) return null
    if ('value' in cell) return cell.value

    if (this.evaluating.has(name)) return CIRCULAR
    this.evaluating.add(name)
    try {
      // `engine.build` costs roughly as much as several interpreted runs, so a
      // sheet that is read once would lose by building everything. Interpret the
      // first evaluation and build only cells that come back — recalculation after
      // a `set` is what makes a cell worth compiling. The built form outlives
      // invalidation, since `set` changes values but not this cell's logic.
      let result
      if (cell.run) {
        result = cell.run(this.data)
      } else {
        result = this.engine.run(cell.logic, this.data)
        // Compile a cell once it proves hot — unless interpreted mode is on, where
        // we never codegen because strict CSP forbids the compiler's eval().
        if (!this.interpreted && cell.evaluations++) cell.run = this.engine.build(cell.logic)
      }
      const value = result === undefined ? null : result
      this.cache.set(name, value)
      return value
    } finally {
      this.evaluating.delete(name)
    }
  }

  // -------------------------------------------------------------------------
  // Dependency tracking
  // -------------------------------------------------------------------------

  /** Record what `name` reads, so a change to any of it can find `name` again. */
  index (name, precedents) {
    for (const precedent of precedents.cells) {
      let readers = this.dependents.get(precedent)
      if (!readers) this.dependents.set(precedent, (readers = new Set()))
      readers.add(name)
    }
    for (const rect of precedents.rects) {
      this.rangeDependents.push({ name, ...rect })
    }
    if (precedents.dynamic) this.dynamic.add(name)
  }

  /**
   * Drop everything recorded for `name`, before its formula is replaced.
   *
   * Driven by the cell's own recorded precedents rather than by sweeping the index:
   * sweeping made every `set` cost O(size of sheet), which on a 10,000-cell sheet was
   * 72 µs per write.
   */
  unindex (name) {
    const cell = this.cells.get(name)
    const precedents = cell && cell.precedents
    if (!precedents) return

    for (const precedent of precedents.cells) {
      const readers = this.dependents.get(precedent)
      if (readers) {
        readers.delete(name)
        if (!readers.size) this.dependents.delete(precedent)
      }
    }
    if (precedents.rects.length) {
      const kept = []
      for (let i = 0; i < this.rangeDependents.length; i++) {
        if (this.rangeDependents[i].name !== name) kept.push(this.rangeDependents[i])
      }
      this.rangeDependents = kept
    }
    if (precedents.dynamic) this.dynamic.delete(name)
  }

  /**
   * Drop `name` and everything downstream of it from the memo.
   *
   * A breadth-first sweep with a visited set, so a circular reference terminates
   * rather than looping. Cells are invalidated by *name*, which means a cell being
   * created for the first time correctly dirties any range that covers it.
   */
  invalidate (name) {
    if (!this.incremental) {
      this.cache.clear()
      return
    }

    const queue = [name]
    const seen = new Set(queue)

    // Anything we could not analyse has to be assumed affected.
    for (const unanalysable of this.dynamic) {
      if (!seen.has(unanalysable)) { seen.add(unanalysable); queue.push(unanalysable) }
    }

    while (queue.length) {
      const current = queue.pop()
      this.cache.delete(current)

      const readers = this.dependents.get(current)
      if (readers) {
        for (const reader of readers) {
          if (!seen.has(reader)) { seen.add(reader); queue.push(reader) }
        }
      }

      for (const entry of this.rangeDependents) {
        if (seen.has(entry.name)) continue
        if (rectContains(entry, current)) { seen.add(entry.name); queue.push(entry.name) }
      }
    }
  }

  /** The cells that read `name`, directly or through a range. For inspection. */
  dependentsOf (name) {
    const direct = this.dependents.get(name)
    const readers = new Set(direct || [])
    for (const entry of this.rangeDependents) {
      if (rectContains(entry, name)) readers.add(entry.name)
    }
    return readers
  }

  // -------------------------------------------------------------------------

  /** The compiled JSON Logic for a formula cell, or null for a literal. */
  logicFor (name) {
    const cell = this.cells.get(name)
    return cell && 'logic' in cell ? cell.logic : null
  }

  /** Every cell, evaluated. Errors appear as their Excel token, e.g. `#DIV/0!`. */
  toJSON () {
    const output = {}
    for (const name of this.cells.keys()) {
      const value = this.get(name)
      output[name] = isError(value) ? value.code : value
    }
    return output
  }
}

export { FormulaSyntaxError, ERROR }
