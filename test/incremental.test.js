import test from 'node:test'
import assert from 'node:assert/strict'
import { Sheet } from '../src/index.js'
import { precedentsOf } from '../src/precedents.js'
import { compile } from '../src/compile.js'

// ---------------------------------------------------------------------------
// Precedent extraction
// ---------------------------------------------------------------------------

const precedents = (formula) => precedentsOf(compile(formula))

test('direct references are found', () => {
  assert.deepEqual([...precedents('=A1 + B2').cells].sort(), ['A1', 'B2'])
  assert.deepEqual([...precedents('=taxRate * total').cells].sort(), ['taxRate', 'total'])
})

test('a deep path depends on its root', () => {
  assert.deepEqual([...precedents('=a.b.c').cells], ['a'])
  assert.deepEqual([...precedents('=Sheet1!A1').cells], ['Sheet1'])
})

test('ranges are kept as rectangles, not expanded', () => {
  const found = precedents('=SUM(A1:E200)')
  assert.equal(found.cells.size, 0)
  assert.deepEqual(found.rects, [{ top: 1, bottom: 200, left: 1, right: 5 }])
})

test('a scoped range depends on the scope root', () => {
  const found = precedents('=SUM(Sheet1!A1:B2)')
  assert.deepEqual([...found.cells], ['Sheet1'])
  assert.deepEqual(found.rects, [])
})

test('nested formulas are walked through', () => {
  const found = precedents('=IF(A1 > 0, SUM(B1:B9), C1 & D1)')
  assert.deepEqual([...found.cells].sort(), ['A1', 'C1', 'D1'])
  assert.deepEqual(found.rects, [{ top: 1, bottom: 9, left: 2, right: 2 }])
})

test('array constants hold no references', () => {
  const found = precedents('=SUM({1,2,3})')
  assert.equal(found.cells.size, 0)
  assert.equal(found.dynamic, false)
})

test('literals have no precedents and are not dynamic', () => {
  const found = precedents('=1 + 2 * 3')
  assert.equal(found.cells.size, 0)
  assert.equal(found.rects.length, 0)
  assert.equal(found.dynamic, false)
})

test('unresolvable references fall back to dynamic', () => {
  // Hand-written logic the grammar would never emit.
  assert.equal(precedentsOf({ val: { cat: ['A', 1] } }).dynamic, true)
  assert.equal(precedentsOf({ val: [[1], 'x'] }).dynamic, true)
  assert.equal(precedentsOf({ RANGE: [{ val: 'x' }, 'B2'] }).dynamic, true)
})

// ---------------------------------------------------------------------------
// Invalidation behaviour
// ---------------------------------------------------------------------------

test('only the affected subgraph is invalidated', () => {
  const sheet = new Sheet({ A1: 1, A2: '=A1+1', A3: '=A2+1', B1: 5, B2: '=B1*2' })
  sheet.toJSON() // populate the memo

  sheet.set('A1', 10)
  // The A-chain is dirty...
  assert.equal(sheet.cache.has('A2'), false)
  assert.equal(sheet.cache.has('A3'), false)
  // ...and the unrelated B-chain is untouched.
  assert.equal(sheet.cache.get('B2'), 10)
  assert.equal(sheet.get('A3'), 12)
})

test('a range dependent is invalidated by any cell inside it', () => {
  const sheet = new Sheet({ A1: 1, A2: 2, A3: 3, B1: '=SUM(A1:A3)', C1: '=99' })
  assert.equal(sheet.get('B1'), 6)

  sheet.set('A2', 20)
  assert.equal(sheet.cache.has('B1'), false)
  assert.equal(sheet.get('B1'), 24)

  // A cell outside the rectangle leaves it alone.
  sheet.get('B1')
  sheet.set('A9', 1000)
  assert.equal(sheet.cache.get('B1'), 24)
})

test('creating a cell inside an existing range dirties that range', () => {
  const sheet = new Sheet({ A1: 1, A3: 3, B1: '=SUM(A1:A5)' })
  assert.equal(sheet.get('B1'), 4)
  sheet.set('A2', 10) // did not exist when B1 was written
  assert.equal(sheet.get('B1'), 14)
})

test('deleting a cell dirties its dependents', () => {
  const sheet = new Sheet({ A1: 5, B1: '=A1*2', C1: '=SUM(A1:A3)' })
  assert.equal(sheet.get('B1'), 10)
  assert.equal(sheet.get('C1'), 5)
  sheet.delete('A1')
  assert.equal(sheet.get('B1'), 0)
  assert.equal(sheet.get('C1'), 0)
})

test('replacing a formula drops its old dependencies', () => {
  const sheet = new Sheet({ A1: 1, B1: 100, C1: '=A1' })
  assert.equal(sheet.get('C1'), 1)

  sheet.set('C1', '=B1')
  assert.equal(sheet.get('C1'), 100)

  // C1 no longer reads A1, so changing A1 must not dirty it.
  sheet.set('A1', 999)
  assert.equal(sheet.cache.get('C1'), 100)
  assert.equal(sheet.dependentsOf('A1').has('C1'), false)
})

test('a dynamic cell is invalidated by any change', () => {
  const sheet = new Sheet({ A1: 1, B1: 2 })
  // Inject logic the grammar cannot produce, to exercise the fallback.
  sheet.cells.set('X1', { logic: { val: [[1], 'A1'] }, run: null, evaluations: 0 })
  sheet.index('X1', precedentsOf({ val: [[1], 'A1'] }))
  assert.equal(sheet.dynamic.has('X1'), true)

  sheet.cache.set('X1', 'stale')
  sheet.set('B1', 3)
  assert.equal(sheet.cache.has('X1'), false)
})

test('dependentsOf reports both direct and range readers', () => {
  const sheet = new Sheet({ A2: 1, B1: '=A2*2', C1: '=SUM(A1:A5)', D1: '=B1' })
  assert.deepEqual([...sheet.dependentsOf('A2')].sort(), ['B1', 'C1'])
  assert.deepEqual([...sheet.dependentsOf('B1')], ['D1'])
})

test('a cycle terminates invalidation instead of looping', () => {
  const sheet = new Sheet({ A1: '=B1', B1: '=A1', C1: 1 })
  sheet.toJSON()
  sheet.set('C1', 2) // must return
  assert.equal(sheet.get('A1').code, '#CIRCULAR!')
})

// ---------------------------------------------------------------------------
// Differential test: the incremental index against the clear-everything oracle
// ---------------------------------------------------------------------------

/** Deterministic PRNG, so a failure can be reproduced from its seed. */
function mulberry32 (seed) {
  return function () {
    seed |= 0; seed = (seed + 0x6D2B79F5) | 0
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed)
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296
  }
}

const CELL_COUNT = 30

/**
 * A formula for cell A{index} that only ever refers to lower-numbered cells, so the
 * generated sheet stays acyclic — a cycle's value depends on which cell is entered
 * first, which would make the comparison order-dependent rather than wrong.
 */
function randomFormula (random, index) {
  const below = () => 1 + Math.floor(random() * (index - 1))
  const shape = Math.floor(random() * 6)
  switch (shape) {
    case 0: return `=A${below()} + A${below()}`
    case 1: return `=A${below()} * 2`
    case 2: return `=SUM(A1:A${below()})`
    case 3: return `=IF(A${below()} > 5, A${below()}, 0)`
    case 4: return `=AVERAGE(A1:A${below()}) + A${below()}`
    default: return `=IFERROR(A${below()} / A${below()}, -1)`
  }
}

function randomValue (random) {
  return Math.floor(random() * 20) - 5
}

test('incremental invalidation matches the clear-everything oracle', () => {
  for (let seed = 1; seed <= 40; seed++) {
    const random = mulberry32(seed)

    const initial = { A1: randomValue(random) }
    for (let i = 2; i <= CELL_COUNT; i++) {
      initial[`A${i}`] = random() < 0.35 ? randomValue(random) : randomFormula(random, i)
    }

    const incremental = new Sheet(initial)
    const oracle = new Sheet(initial, { incremental: false })

    assert.deepEqual(incremental.toJSON(), oracle.toJSON(), `seed ${seed}: initial state`)

    for (let step = 0; step < 60; step++) {
      const index = 1 + Math.floor(random() * CELL_COUNT)
      const name = `A${index}`
      const action = random()

      let description
      if (action < 0.45) {
        const value = randomValue(random)
        description = `set ${name} = ${value}`
        incremental.set(name, value)
        oracle.set(name, value)
      } else if (action < 0.85 && index > 1) {
        const formula = randomFormula(random, index)
        description = `set ${name} = ${formula}`
        incremental.set(name, formula)
        oracle.set(name, formula)
      } else {
        description = `delete ${name}`
        incremental.delete(name)
        oracle.delete(name)
      }

      // Read a few cells before comparing, so the incremental sheet carries a
      // partially populated memo into the next mutation — that is where a stale
      // entry would survive.
      for (let read = 0; read < 5; read++) {
        const target = `A${1 + Math.floor(random() * CELL_COUNT)}`
        incremental.get(target)
        oracle.get(target)
      }

      assert.deepEqual(
        incremental.toJSON(),
        oracle.toJSON(),
        `seed ${seed}, step ${step}: after ${description}`
      )
    }
  }
})
