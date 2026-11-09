import ast
import json
import subprocess
import unittest

from py2ts_steps import S
from py2ts_tests import T


class ExceptionTranslationTests(unittest.TestCase):
    def execute(self, translator, source, assertions):
        function = ast.parse(source).body[0]
        generated = translator().func_expr(function, 0)
        program = """
import assert from 'node:assert/strict';
import ts from 'typescript';
const $isinstance = (error, cls) => Array.isArray(cls)
  ? cls.some(type => error instanceof type) : error instanceof cls;
const $call = (fn, ...args) => fn(...args);
class ValueError extends Error {}
const check = eval(ts.transpileModule(%s, {
  compilerOptions: { target: ts.ScriptTarget.ES2022 }
}).outputText);
%s
""" % (json.dumps("(" + generated + ")"), assertions)
        result = subprocess.run(
            ["node", "--input-type=module", "-e", program],
            capture_output=True, text=True,
        )
        self.assertEqual(result.returncode, 0, result.stderr)

    def test_dynamic_exception_rejects_success_and_wrong_type(self):
        source = '''
def check(operation, expected):
    try:
        operation()
        raise AssertionError("did not raise")
    except expected:
        return "caught"
'''
        assertions = '''
assert.equal(check(() => { throw new TypeError(); }, TypeError), 'caught');
assert.throws(() => check(() => {}, TypeError), /did not raise/);
const wrong = new ValueError('wrong type');
assert.throws(() => check(() => { throw wrong; }, TypeError), error => error === wrong);
'''
        for translator in (T, S):
            with self.subTest(translator=translator.__name__):
                self.execute(translator, source, assertions)

    def test_ordered_handlers_bind_exception_and_run_finally(self):
        source = '''
def check(operation, cleanup):
    try:
        operation()
    except TypeError as error:
        return error
    except (ValueError, RangeError) as error:
        return error
    except Exception:
        raise
    finally:
        cleanup()
'''
        assertions = '''
let cleanups = 0;
const cleanup = () => { cleanups++; };
assert.equal(check(() => {}, cleanup), undefined);
for (const error of [new TypeError(), new ValueError(), new RangeError()]) {
  assert.equal(check(() => { throw error; }, cleanup), error);
}
const unhandled = new Error('unhandled');
assert.throws(() => check(() => { throw unhandled; }, cleanup), error => error === unhandled);
assert.equal(cleanups, 5);
'''
        for translator in (T, S):
            with self.subTest(translator=translator.__name__):
                self.execute(translator, source, assertions)


if __name__ == "__main__":
    unittest.main()
