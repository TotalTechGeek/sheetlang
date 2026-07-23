import test from 'node:test'
import assert from 'node:assert/strict'
import { Sheet, isError, FormulaSyntaxError } from '../src/index.js'

test('literal cells come back untouched', () => {
  const sheet = new Sheet({ A1: 5, A2: 'text', A3: true })
  assert.equal(sheet.get('A1'), 5)
  assert.equal(sheet.get('A2'), 'text')
  assert.equal(sheet.get('A3'), true)
})

test('an unset cell reads as blank', () => {
  assert.equal(new Sheet().get('Z9'), null)
})

test('formulas resolve their dependencies on demand', () => {
  const sheet = new Sheet({
    A1: 10,
    A2: 20,
    A3: '=A1+A2',
    A4: '=A3*2'
  })
  assert.equal(sheet.get('A4'), 60)
})

test('dependency order does not matter', () => {
  const sheet = new Sheet({ C1: '=B1+1', B1: '=A1+1', A1: 1 })
  assert.equal(sheet.get('C1'), 3)
})

test('ranges see computed values, not formulas', () => {
  const sheet = new Sheet({
    A1: 1,
    A2: '=A1*2',
    A3: '=A2*2',
    B1: '=SUM(A1:A3)'
  })
  assert.equal(sheet.get('B1'), 7)
})

test('circular references are detected rather than hanging', () => {
  const sheet = new Sheet({ A1: '=B1', B1: '=A1' })
  const result = sheet.get('A1')
  assert.ok(isError(result))
  assert.equal(result.code, '#CIRCULAR!')
})

test('a cell referring to itself is circular too', () => {
  const sheet = new Sheet({ A1: '=A1+1' })
  assert.equal(sheet.get('A1').code, '#CIRCULAR!')
})

test('setting a cell recalculates dependents', () => {
  const sheet = new Sheet({ A1: 2, A2: '=A1*10' })
  assert.equal(sheet.get('A2'), 20)
  sheet.set('A1', 3)
  assert.equal(sheet.get('A2'), 30)
})

test('deleting a cell leaves dependents seeing a blank', () => {
  const sheet = new Sheet({ A1: 2, A2: '=A1+1' })
  assert.equal(sheet.get('A2'), 3)
  sheet.delete('A1')
  assert.equal(sheet.get('A2'), 1)
})

test('syntax errors surface when the cell is written', () => {
  const sheet = new Sheet()
  assert.throws(() => sheet.set('A1', '=SUM('), FormulaSyntaxError)
})

test('errors flow through the sheet as values', () => {
  const sheet = new Sheet({ A1: 0, A2: '=10/A1', A3: '=IFERROR(A2, "n/a")' })
  assert.equal(sheet.get('A2').code, '#DIV/0!')
  assert.equal(sheet.get('A3'), 'n/a')
})

test('the compiled logic is inspectable', () => {
  const sheet = new Sheet({ A1: 1, A2: '=A1*2' })
  assert.deepEqual(sheet.logicFor('A2'), { '*': [{ val: 'A1' }, 2] })
  assert.equal(sheet.logicFor('A1'), null)
})

test('toJSON renders errors as their Excel token', () => {
  const sheet = new Sheet({ A1: 1, A2: '=A1+1', A3: '=1/0' })
  assert.deepEqual(sheet.toJSON(), { A1: 1, A2: 2, A3: '#DIV/0!' })
})

test('a worked example: an invoice', () => {
  const sheet = new Sheet({
    B2: 3, C2: 19.99,
    B3: 1, C3: 249.0,
    B4: 5, C4: 4.5,
    D2: '=B2*C2',
    D3: '=B3*C3',
    D4: '=B4*C4',
    D6: '=SUM(D2:D4)',
    D7: '=ROUND(D6 * VAT, 2)',
    D8: '=D6+D7',
    VAT: 0.2
  })
  assert.equal(sheet.get('D6'), 331.47)
  assert.equal(sheet.get('D7'), 66.29)
  assert.equal(sheet.get('D8').toFixed(2), '397.76')
})
