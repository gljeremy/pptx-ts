"""Translate python-pptx's behave step modules (features/steps/*.py) into TypeScript.

Usage: py -I scripts/py2ts_steps.py <in.py> <out.ts>

Uses the expression and statement translation of scripts/py2ts_tests.py. Each
`@given/@when/@then("pattern")` function becomes `step("given", "pattern", fn)` for the
Gherkin runner in test/features/runner.ts. `eval(s)` and `literal_eval(s)` become
`$pyEval(s, $scope)`, where `$scope` holds the module's imported names (what python's `eval`
sees). `from helpers import ...` maps to the hand-written test/features/steps/helpers.ts.
"""
import ast
import importlib.util
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
spec = importlib.util.spec_from_file_location("py2ts_tests", os.path.join(HERE, "py2ts_tests.py"))
tt = importlib.util.module_from_spec(spec)
spec.loader.exec_module(tt)

STEP_KINDS = ("given", "when", "then", "step")


class S(tt.T):
    def e_Call(self, n):
        f = n.func
        if isinstance(f, ast.Name) and f.id in ("eval", "literal_eval") and not self.is_local(f.id):
            scope = "$scope" if f.id == "eval" else "{}"
            return f"$pyEval({self.expr(n.args[0])}, {scope})"
        if isinstance(f, ast.Attribute) and isinstance(f.value, ast.Name) and self.module_aliases.get(f.value.id) == "os":
            return f"$os.{f.attr}({self.args_js(n.args, [])})"
        return super().e_Call(n)


def step_info(fn):
    """[(kind, pattern)] in registration order (decorators apply bottom-up)."""
    out = []
    for d in fn.decorator_list:
        if isinstance(d, ast.Call) and isinstance(d.func, ast.Name) and d.func.id in STEP_KINDS:
            out.append((d.func.id, d.args[0].value))
    return out[::-1]


def main():
    src, dest = sys.argv[1], sys.argv[2]
    tree = ast.parse(open(src, encoding="utf-8").read())
    t = S()
    for s in tree.body:
        if isinstance(s, (ast.Import, ast.ImportFrom)):
            if isinstance(s, ast.ImportFrom) and s.module in ("behave", "behave.runner", "helpers"):
                continue
            t.stmt(s, 0)
        elif isinstance(s, ast.ClassDef):
            t.classes.add(s.name)
        elif isinstance(s, ast.FunctionDef) and not step_info(s):
            t.local_funcs.add(s.name)
    helpers = sorted(set(
        a.asname or a.name
        for s in ast.walk(tree)
        if isinstance(s, ast.ImportFrom) and s.module == "helpers"
        for a in s.names
    ))
    lines = []
    for s in tree.body:
        if isinstance(s, (ast.Import, ast.ImportFrom)):
            continue
        if isinstance(s, ast.FunctionDef):
            info = step_info(s)
            params = [x.arg for x in s.args.args]
            plist = t.params(s, skip_self=False)
            body = t.func_body(s, 0, params)
            lines.append(f"function {s.name}({', '.join(plist)}): any {{")
            lines += body
            lines.append("}")
            for kind, pattern in info:
                lines.append(f"step({json.dumps(kind)}, {tt.json_str(pattern)}, {json.dumps(params[1:])}, {s.name});")
        elif isinstance(s, ast.ClassDef):
            lines += t.helper_class(s, 0)
        else:
            lines += t.stmt(s, 0)
        lines.append("")

    test_root = os.path.abspath(os.path.join(HERE, "..", "test"))
    out_dir = os.path.dirname(os.path.abspath(dest))
    rel_test = os.path.relpath(test_root, out_dir).replace("\\", "/")
    rel_src = os.path.relpath(os.path.join(test_root, "..", "src"), out_dir).replace("\\", "/")
    header = [
        f"// Machine-translated from python-pptx {src.replace(chr(92), '/').split('python-pptx/')[-1]} by",
        "// scripts/py2ts_steps.py. Python semantics come from test/pyshim.ts.",
        "/* eslint-disable */",
        "// @ts-nocheck",
        f'import * as $py from "{rel_test}/pyshim.js";',
        "__SHIM_IMPORTS__",
        f'import {{ AttributeError, IndexError, KeyError, NotImplementedError, ValueError }} from "{rel_src}/errors.js";',
        f'import {{ step }} from "../runner.js";',
    ]
    if helpers:
        header.append(f'import {{ {", ".join(helpers)} }} from "./helpers.js";')
    t.pkg_parts = []
    imports = tt.import_lines(t, rel_src, rel_test)
    # step modules aren't patched by mock helpers: import the bindings directly
    imports = [l for l in imports if not l.startswith("//") and not l.endswith("$imported;")]
    imports = [re.sub(r"\b(\w+) as \1\$imported\b", r"\1", l).replace("$imported", "") for l in imports]
    header += imports
    header.append("void $py;")
    header.append("")
    # python eval() sees the module globals; the expressions only use pptx names
    scope_names = [n for n in sorted(t.imports) if t.imports[n][0].startswith("pptx")]
    lines.append("// the names python's eval() sees in this module")
    lines.append("const $scope = { " + ", ".join(tt.camel(n) if n in t.camel_names else n for n in scope_names) + " };")
    body_text = "\n".join(lines)
    shim_names = r"(?<![\w$.])(\$[A-Za-z_]\w*|BytesIO|pytest)(?![\w$])"
    used = sorted(set(re.findall(shim_names, body_text)) - {"$py", "$scope"})
    header = [f'import {{ {", ".join(used)} }} from "{rel_test}/pyshim.js";' if h == "__SHIM_IMPORTS__" else h for h in header]
    text = "\n".join(header + lines) + "\n"
    fixups_path = os.path.join(test_root, "fixups.json")
    if os.path.exists(fixups_path):
        rel_dest = os.path.relpath(os.path.abspath(dest), os.path.dirname(test_root)).replace("\\", "/")
        for find, repl, *_ in json.load(open(fixups_path, encoding="utf-8")).get(rel_dest, []):
            if find not in text:
                print(f"warning: fixup not applied in {rel_dest}: {find[:60]!r}", file=sys.stderr)
            text = text.replace(find, repl)
    os.makedirs(out_dir, exist_ok=True)
    open(dest, "w", encoding="utf-8").write(text)


if __name__ == "__main__":
    main()
