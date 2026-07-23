import test from 'node:test'
import assert from 'node:assert/strict'
import { build, evaluate, Sheet } from '../src/index.js'

/**
 * Under a strict Content-Security-Policy (no `unsafe-eval`) the engine's compiler
 * cannot run, because it materialises functions with `eval`. Interpreted mode must
 * never reach that path.
 *
 * We cannot install a real CSP in Node, so we make `eval` and the `Function`
 * constructor throw and assert the interpreted paths survive — and, as a control,
 * that the compiled path genuinely would have failed.
 */
function withEvalBlocked (run) {
  const realEval = globalThis.eval
  const realFunction = globalThis.Function
  globalThis.eval = () => { throw new EvalError('CSP: eval blocked') }
  globalThis.Function = new Proxy(realFunction, {
    construct () { throw new EvalError('CSP: Function constructor blocked') },
    apply () { throw new EvalError('CSP: Function constructor blocked') }
  })
  try {
    return run()
  } finally {
    globalThis.eval = realEval
    globalThis.Function = realFunction
  }
}

test('the control: the compiled build() path really does need eval', () => {
  withEvalBlocked(() => {
    assert.throws(() => build('=1+1')(), /CSP/)
  })
})

test('evaluate() runs under a strict CSP', () => {
  withEvalBlocked(() => {
    assert.equal(evaluate('=SUM(A1:A3) * 2', { A1: 1, A2: 2, A3: 3 }), 12)
    assert.equal(evaluate('=IF(x>0, ROUND(x*1.2, 2), 0)', { x: 10 }), 12)
  })
})

test('build({ interpreted: true }) runs under a strict CSP', () => {
  withEvalBlocked(() => {
    const priced = build('=ROUND(qty * price, 2)', { interpreted: true })
    assert.equal(priced({ qty: 3, price: 19.99 }), 59.97)
    assert.equal(priced({ qty: 2, price: 5 }), 10)
  })
})

test('an interpreted Sheet never compiles, even when a cell turns hot', () => {
  withEvalBlocked(() => {
    const sheet = new Sheet({ A1: 2, A2: '=A1*10', A3: '=SUM(A1:A2)' }, { interpreted: true })
    assert.equal(sheet.get('A2'), 20)
    // Repeated reads across edits are exactly what would trigger engine.build().
    sheet.set('A1', 5); assert.equal(sheet.get('A2'), 50)
    sheet.set('A1', 7); assert.equal(sheet.get('A2'), 70)
    sheet.set('A1', 9); assert.equal(sheet.get('A2'), 90)
    assert.equal(sheet.get('A3'), 99)
  })
})

test('interpreted mode agrees with compiled mode', () => {
  const formula = '=IF(AND(a>1, b<10), ROUND(SUM(a, b) * rate, 2), -1)'
  const data = { a: 3, b: 4, rate: 1.5 }
  assert.equal(build(formula, { interpreted: true })(data), build(formula)(data))
})
