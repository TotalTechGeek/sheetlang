/**
 * A1-style reference maths. Ranges are expanded at runtime rather than at compile
 * time so that `RANGE` stays a normal JSON Logic operation over whatever data you
 * hand the engine.
 */

import { REF } from './errors.js'

const A1_PATTERN = /^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)$/

/** `A` -> 1, `Z` -> 26, `AA` -> 27 */
export function columnToIndex (letters) {
  let index = 0
  const upper = letters.toUpperCase()
  for (let i = 0; i < upper.length; i++) {
    index = index * 26 + (upper.charCodeAt(i) - 64)
  }
  return index
}

const COLUMN_NAMES = []

/** 1 -> `A`, 27 -> `AA` */
export function indexToColumn (index) {
  const cached = COLUMN_NAMES[index]
  if (cached !== undefined) return cached

  let letters = ''
  let n = index
  while (n > 0) {
    const remainder = (n - 1) % 26
    letters = String.fromCharCode(65 + remainder) + letters
    n = Math.floor((n - 1) / 26)
  }
  COLUMN_NAMES[index] = letters
  return letters
}

/** Parse `A1`/`$A$1` into `{ col, row }` (both 1-based), or null if it isn't one. */
export function parseA1 (ref) {
  if (typeof ref !== 'string') return null
  const match = A1_PATTERN.exec(ref.trim())
  if (!match) return null
  return { col: columnToIndex(match[1]), row: Number(match[2]) }
}

export function isA1 (ref) {
  return parseA1(ref) !== null
}

/**
 * The names in a range never change, but `RANGE` is re-evaluated on every run, so
 * the expansion is memoised. Without this a `SUM(A1:E200)` rebuilds a thousand
 * strings per evaluation, which dominates the cost of the whole formula.
 */
const RANGE_CACHE = new Map()
const RANGE_CACHE_LIMIT = 4096

/**
 * Every cell name in the rectangle bounded by two corners, as rows of names.
 * The corners may be given in any order, as Excel normalises `B3:A1`.
 *
 * The returned grid is shared and must not be mutated by callers.
 * @returns {readonly string[][]} row-major grid of cell names
 */
export function expandRange (start, end) {
  const key = `${start}:${end}`
  const cached = RANGE_CACHE.get(key)
  if (cached !== undefined) return cached

  const a = parseA1(start)
  const b = parseA1(end)
  if (!a || !b) throw REF

  const top = Math.min(a.row, b.row)
  const bottom = Math.max(a.row, b.row)
  const left = Math.min(a.col, b.col)
  const right = Math.max(a.col, b.col)

  const grid = new Array(bottom - top + 1)
  for (let row = top; row <= bottom; row++) {
    const cells = new Array(right - left + 1)
    for (let col = left; col <= right; col++) cells[col - left] = indexToColumn(col) + row
    grid[row - top] = cells
  }

  // Frozen because the grid is shared with every caller, including the public
  // `expandRange` export.
  for (const row of grid) Object.freeze(row)
  Object.freeze(grid)

  // A crude bound: formulas reference a bounded set of ranges in practice, so
  // dropping everything on overflow is cheaper than tracking recency.
  if (RANGE_CACHE.size >= RANGE_CACHE_LIMIT) RANGE_CACHE.clear()
  RANGE_CACHE.set(key, grid)
  return grid
}

/** Walk a path of keys into an object, returning undefined on any missing link. */
export function resolvePath (root, path) {
  let current = root
  for (const key of path) {
    if (current === null || current === undefined) return undefined
    current = current[key]
  }
  return current
}
