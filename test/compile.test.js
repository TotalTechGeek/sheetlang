import test from 'node:test'
import assert from 'node:assert/strict'
import { compile, FormulaSyntaxError } from '../src/index.js'

test('operators map onto native JSON Logic names', () => {
  assert.deepEqual(compile('=1+2'), { '+': [1, 2] })
  assert.deepEqual(compile('=1-2'), { '-': [1, 2] })
  assert.deepEqual(compile('=1*2'), { '*': [1, 2] })
  assert.deepEqual(compile('=1/2'), { '/': [1, 2] })
  assert.deepEqual(compile('=1=2'), { '==': [1, 2] })
  assert.deepEqual(compile('=1<>2'), { '!=': [1, 2] })
  assert.deepEqual(compile('=1<=2'), { '<=': [1, 2] })
})

test('precedence follows Excel', () => {
  assert.deepEqual(compile('=1+2*3'), { '+': [1, { '*': [2, 3] }] })
  assert.deepEqual(compile('=(1+2)*3'), { '*': [{ '+': [1, 2] }, 3] })
  // Unary minus binds tighter than ^
  assert.deepEqual(compile('=-2^2'), { POWER: [{ '-': [2] }, 2] })
  // ^ is left-associative in Excel
  assert.deepEqual(compile('=2^3^2'), { POWER: [{ POWER: [2, 3] }, 2] })
  // Comparison is looser than concatenation
  assert.deepEqual(compile('="a"&"b"="ab"'), { '==': [{ CONCAT: ['a', 'b'] }, 'ab'] })
})

test('chained & collapses into one CONCAT', () => {
  assert.deepEqual(compile('="a"&"b"&"c"'), { CONCAT: ['a', 'b', 'c'] })
})

test('the leading = is optional', () => {
  assert.deepEqual(compile('1+1'), compile('=1+1'))
})

test('percent is a postfix division', () => {
  assert.deepEqual(compile('=50%'), { '/': [50, 100] })
  assert.deepEqual(compile('=A1*10%'), { '*': [{ val: 'A1' }, { '/': [10, 100] }] })
})

test('single-segment references use the unary val form', () => {
  assert.deepEqual(compile('=revenue'), { val: 'revenue' })
})

test('deep paths use the array val form', () => {
  assert.deepEqual(compile('=a.b.c'), { val: ['a', 'b', 'c'] })
  assert.deepEqual(compile('=items[0].name'), { val: ['items', 0, 'name'] })
  assert.deepEqual(compile('=data["odd key"]'), { val: ['data', 'odd key'] })
})

test('bracketed names carry spaces', () => {
  assert.deepEqual(compile('=[Total Sales]'), { val: 'Total Sales' })
})

test('scope prefixes become leading path segments', () => {
  assert.deepEqual(compile('=Sheet1!A1'), { val: ['Sheet1', 'A1'] })
  assert.deepEqual(compile("='My Sheet'!A1"), { val: ['My Sheet', 'A1'] })
})

test('ranges compile to a RANGE operation', () => {
  assert.deepEqual(compile('=A1:B3'), { RANGE: ['A1', 'B3'] })
  assert.deepEqual(compile('=$A$1:$B$3'), { RANGE: ['A1', 'B3'] })
  assert.deepEqual(compile('=Sheet1!A1:B3'), { RANGE: ['A1', 'B3', { preserve: ['Sheet1'] }] })
})

test('anchored single cells normalise, plain ones stay verbatim', () => {
  assert.deepEqual(compile('=$a$1'), { val: 'A1' })
  assert.deepEqual(compile('=a1'), { val: 'a1' })
})

test('function names are case-insensitive and normalise to uppercase', () => {
  assert.deepEqual(compile('=sum(1,2)'), { SUM: [1, 2] })
  assert.deepEqual(compile('=Sum(1,2)'), { SUM: [1, 2] })
})

test('omitted arguments become null', () => {
  assert.deepEqual(compile('=IF(A1,,2)'), { IF: [{ val: 'A1' }, null, 2] })
  assert.deepEqual(compile('=TODAY()'), { TODAY: [] })
})

test('literals', () => {
  assert.equal(compile('=1.5e3'), 1500)
  assert.equal(compile('=.5'), 0.5)
  assert.equal(compile('="he said ""hi"""'), 'he said "hi"')
  assert.equal(compile('=TRUE'), true)
  assert.equal(compile('=false'), false)
  assert.deepEqual(compile('=#REF!'), { ERRORVALUE: ['#REF!'] })
})

test('array constants are preserved rather than evaluated', () => {
  assert.deepEqual(compile('={1,2,3}'), { preserve: [1, 2, 3] })
  assert.deepEqual(compile('={1,2;3,4}'), { preserve: [[1, 2], [3, 4]] })
  assert.deepEqual(compile('={-1,"a",TRUE}'), { preserve: [-1, 'a', true] })
})

test('whitespace and newlines are ignored', () => {
  assert.deepEqual(compile('=  SUM( 1 ,\n  2 )  '), { SUM: [1, 2] })
})

test('unknown functions are rejected at compile time', () => {
  assert.throws(() => compile('=NOPE(1)'), FormulaSyntaxError)
  assert.throws(() => compile('=NOPE(1)'), /Unknown function NOPE/)
})

test('the function whitelist can be replaced or disabled', () => {
  assert.deepEqual(compile('=NOPE(1)', { functions: null }), { NOPE: [1] })
  assert.deepEqual(compile('=nope(1)', { functions: ['NOPE'] }), { NOPE: [1] })
  assert.throws(() => compile('=SUM(1)', { functions: ['NOPE'] }), /Unknown function SUM/)
})

test('syntax errors report a location', () => {
  assert.throws(() => compile('=1 +'), FormulaSyntaxError)
  try {
    compile('=1 +')
  } catch (thrown) {
    assert.ok(thrown.location)
    assert.equal(thrown.location.start.line, 1)
    assert.match(thrown.message, /column 5/)
  }
})

test('non-string input is rejected', () => {
  assert.throws(() => compile(42), FormulaSyntaxError)
})

test('unary functions emit the collapsed form', () => {
  // `optimizeUnary` methods take their argument unwrapped, so the grammar does not
  // add a wrapper the engine would only have to strip again.
  assert.deepEqual(compile('=ABS(-3)'), { ABS: { '-': [3] } })
  assert.deepEqual(compile('=LEN("abc")'), { LEN: 'abc' })
  assert.deepEqual(compile('=UPPER(name)'), { UPPER: { val: 'name' } })
  assert.deepEqual(compile('=ISERROR(A1)'), { ISERROR: { val: 'A1' } })
})

test('variadic and optional-argument functions keep the array form', () => {
  assert.deepEqual(compile('=ROUND(2.5)'), { ROUND: [2.5] })
  assert.deepEqual(compile('=LOG(100, 10)'), { LOG: [100, 10] })
  assert.deepEqual(compile('=SUM(1, 2)'), { SUM: [1, 2] })
  assert.deepEqual(compile('=ABS()'), { ABS: [] })
})

test('the unary set can be overridden for a custom engine', () => {
  assert.deepEqual(compile('=ABS(1)', { unaryFunctions: [] }), { ABS: [1] })
  assert.deepEqual(compile('=ROUND(1)', { unaryFunctions: ['ROUND'] }), { ROUND: 1 })
})

test('the [*] projection sugar compiles to a map', () => {
  // `items[*].total` projects `total` over every element via json-logic's map.
  assert.deepEqual(compile('=items[*].total'),
    { map: [{ val: 'items' }, { val: 'total' }] })
  assert.deepEqual(compile('=SUM(items[*].total)'),
    { SUM: [{ map: [{ val: 'items' }, { val: 'total' }] }] })
  // A deeper projection path stays a path inside the mapper.
  assert.deepEqual(compile('=orders[*].customer.city'),
    { map: [{ val: 'orders' }, { val: ['customer', 'city'] }] })
  // `items[*]` with nothing after is just the array.
  assert.deepEqual(compile('=items[*]'), { val: 'items' })
})

test('nested [*] projections nest maps', () => {
  assert.deepEqual(compile('=rows[*].cells[*].v'),
    { map: [{ val: 'rows' }, { map: [{ val: 'cells' }, { val: 'v' }] }] })
})
