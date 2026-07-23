/**
 * Static dependency extraction: which cells does a compiled formula read?
 *
 * This is only possible because the AST is data. The grammar emits literal
 * references — `{ val: 'A1' }`, `{ RANGE: ['A1', 'E200'] }` — so a formula's
 * precedents can be read straight off the tree without evaluating anything.
 *
 * Anything that cannot be resolved statically (a computed reference, a hand-written
 * `val` using the scope-climbing form) marks the whole formula `dynamic`, and the
 * sheet falls back to invalidating it on every change. Correctness before precision.
 */

import { parseA1 } from './refs.js'

/**
 * @typedef {object} Precedents
 * @property {Set<string>} cells Directly referenced names, e.g. `A1`, `taxRate`
 * @property {{top:number,left:number,bottom:number,right:number}[]} rects
 *   Ranges, kept as rectangles rather than expanded — a `SUM(A1:E200)` would
 *   otherwise contribute a thousand index entries on its own.
 * @property {boolean} dynamic True when references could not be resolved statically
 */

/**
 * @param {any} logic A compiled formula
 * @returns {Precedents}
 */
export function precedentsOf (logic) {
  const found = { cells: new Set(), rects: [], dynamic: false }
  walk(logic, found)
  return found
}

function walk (logic, found) {
  if (Array.isArray(logic)) {
    for (const item of logic) walk(item, found)
    return
  }
  if (!logic || typeof logic !== 'object') return

  for (const operator of Object.keys(logic)) {
    const args = logic[operator]

    // `preserve` marks data the engine must not interpret, so it holds no references.
    if (operator === 'preserve') continue

    if (operator === 'val') {
      readVal(args, found)
      continue
    }

    if (operator === 'RANGE') {
      readRange(args, found)
      continue
    }

    walk(args, found)
  }
}

/** `{ val: 'A1' }`, `{ val: ['a','b','c'] }` — the dependency is the first segment. */
function readVal (args, found) {
  if (typeof args === 'string' || typeof args === 'number') {
    found.cells.add(String(args))
    return
  }
  if (Array.isArray(args)) {
    const root = args[0]
    // `[[1], 'x']` is the scope-climbing form: it reads from an outer context, not
    // from this sheet, and cannot be attributed to a cell.
    if (Array.isArray(root)) { found.dynamic = true; return }
    if (typeof root === 'string' || typeof root === 'number') {
      found.cells.add(String(root))
      return
    }
  }
  // A computed reference, e.g. { val: { cat: [...] } }
  found.dynamic = true
  walk(args, found)
}

/** `{ RANGE: [start, end, scope?] }` — recorded as a rectangle. */
function readRange (args, found) {
  if (!Array.isArray(args)) { found.dynamic = true; return }
  const [start, end, scope] = args

  if (scope !== undefined) {
    // A scoped range reads from nested data, so the dependency is that root.
    const path = scope && typeof scope === 'object' ? scope.preserve : undefined
    if (Array.isArray(path) && typeof path[0] === 'string') found.cells.add(path[0])
    else found.dynamic = true
    return
  }

  if (typeof start !== 'string' || typeof end !== 'string') {
    found.dynamic = true
    walk(args, found)
    return
  }

  const a = parseA1(start)
  const b = parseA1(end)
  if (!a || !b) { found.dynamic = true; return }

  found.rects.push({
    top: Math.min(a.row, b.row),
    bottom: Math.max(a.row, b.row),
    left: Math.min(a.col, b.col),
    right: Math.max(a.col, b.col)
  })
}

/** Does a cell name fall inside a rectangle? Non-A1 names never do. */
export function rectContains (rect, name) {
  const position = parseA1(name)
  if (!position) return false
  return position.row >= rect.top && position.row <= rect.bottom &&
    position.col >= rect.left && position.col <= rect.right
}
