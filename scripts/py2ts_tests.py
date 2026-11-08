"""Translate python-pptx pytest modules into vitest TypeScript.

Usage: py -I scripts/py2ts_tests.py <in.py> <out.test.ts>

A porting aid for the style of python-pptx's unit tests (Describe* classes, fixtures,
unittest.mock helpers), not a general Python compiler. Python semantics come from
test/pyshim.ts and test/mock.ts; the pytest structure becomes `pytestClass(...)` calls that
test/pyshim.ts expands into vitest cases. Output is reviewed and fixed by hand where needed.

Naming: attribute and method names, and functions imported from pptx or the test utilities,
are converted to camelCase (`get_or_add_spPr` -> `getOrAddSpPr`); class names, CONSTANTS and
local variables are kept.
"""
import ast
import os
import re
import sys

EXC = {
    "ValueError": "ValueError", "KeyError": "KeyError", "IndexError": "IndexError", "TypeError": "TypeError",
    "AttributeError": "AttributeError", "NotImplementedError": "NotImplementedError",
    "ZeroDivisionError": "ZeroDivisionError", "Exception": "null", "BaseException": "null",
    "InvalidXmlError": "InvalidXmlError", "PackageNotFoundError": "PackageNotFoundError",
    "PythonPptxError": "PythonPptxError", "FileNotFoundError": "FileNotFoundError", "IOError": "null",
    "OSError": "null", "AssertionError": "null", "StopIteration": "null", "UnicodeDecodeError": "null",
    "XMLSyntaxError": "null", "BadZipFile": "null", "RuntimeError": "null",
}

BUILTINS = {
    "len": "$len", "list": "$list", "tuple": "$list", "range": "$range", "str": "$str", "repr": "$repr",
    "sorted": "$sorted", "zip": "$zip", "enumerate": "$enumerate", "any": "$any", "all": "$all", "sum": "$sum",
    "abs": "$abs", "float": "$float", "int": "$int", "bool": "$truthy", "min": "$min", "max": "$max",
    "print": "console.log", "set": "$set_", "frozenset": "$set_", "id": "$id", "getattr": "$getattr",
    "hasattr": "$hasattr", "setattr": "$setattr", "callable": "$callable", "round": "$round", "bytes": "$bytes",
    "iter": "$iter", "next": "$next", "type": "$type", "dict": "$dict", "isinstance": "$isinstance",
    "bytearray": "$bytes", "cast": "$cast", "open": "$open",
}

# test-utility modules whose function names get camelCased
CAMEL_MODULE_PREFIXES = ("pptx", "..unitutil", "unitutil", "tests")

# python methods of builtin types whose JS spelling differs or that collide with src methods
DISPATCH = {
    "append": "$append", "extend": "$extend", "insert": "$insert", "remove": "$remove", "pop": "$pop",
    "index": "$index", "count": "$count", "find": "$find", "replace": "$replace", "copy": "$copy",
    "update": "$update", "keys": "$keys", "values": "$values", "items": "$items", "get": "$dictget",
    "split": "$split", "strip": "$strip", "lstrip": "$lstrip", "rstrip": "$rstrip", "lower": "$lower",
    "upper": "$upper", "startswith": "$startswith", "endswith": "$endswith", "encode": "$encode",
    "decode": "$decode", "splitlines": "$splitlines", "isdigit": "$isdigit", "setdefault": "$setdefault",
    "read": "$read", "clear": "$clear", "sort": "$sort",
}
# dispatch only when the receiver is not obviously a src object
DISPATCH_ALWAYS = {"split", "strip", "lstrip", "rstrip", "lower", "upper", "startswith", "endswith", "encode",
                   "decode", "splitlines", "isdigit", "setdefault", "extend"}


JS_RESERVED = {
    "package", "default", "new", "class", "delete", "function", "var", "let", "in", "this", "with", "yield",
    "interface", "private", "public", "protected", "static", "enum", "export", "import", "super", "switch",
    "case", "typeof", "void", "arguments", "eval", "implements", "instanceof", "const", "do", "catch",
    "finally", "throw", "debugger", "extends", "var", "await", "null", "true", "false",
}


def ident(name):
    """A python local name as a valid JS identifier."""
    return name + "$" if name in JS_RESERVED else name


def camel(name):
    m = re.match(r"^_*", name).group(0)
    rest = name[len(m):]
    if rest.startswith("_") or rest.endswith("__"):
        return name
    return m + re.sub(r"_+([A-Za-z0-9])", lambda x: x.group(1).upper(), rest)


def attr_name(name):
    """camelCase an attribute name unless it's a dunder, a CONSTANT or a ClassName."""
    if name.startswith("__") and name.endswith("__"):
        return name
    core = name.lstrip("_")
    if not core:
        return name
    if core.isupper() or core[0].isupper():
        return name
    return camel(name)


def repr_js(s):
    return json_str(s)


def json_str(s):
    out = '"'
    for ch in s:
        o = ord(ch)
        if ch == "\\":
            out += "\\\\"
        elif ch == '"':
            out += '\\"'
        elif ch == "\n":
            out += "\\n"
        elif ch == "\r":
            out += "\\r"
        elif ch == "\t":
            out += "\\t"
        elif o < 0x20 or 0x7f <= o < 0xa0 or 0x2028 <= o <= 0x2029:
            out += "\\u%04x" % o
        else:
            out += ch
    return out + '"'


def bytes_js(b):
    s = "".join(chr(x) for x in b)
    return f"$b({json_str(s)})"


class T:
    def __init__(self):
        self.camel_names = set()     # imported functions to camelCase
        self.classes = set()         # class names known in this file
        self.test_classes = {}       # Describe* class name -> ClassDef
        self.module_fixtures = {}
        self.scopes = [set()]
        self.local_funcs = set()
        self.imports = {}            # name -> (module, original name)
        self.uses = set()
        self.module_aliases = {}     # local alias -> module path (e.g. "io")

    # ------------------------------------------------------------------ helpers
    def declare(self, name):
        for s in self.scopes:
            if name in s:
                return False
        self.scopes[-1].add(name)
        return True

    def is_local(self, name):
        return any(name in s for s in self.scopes[1:])

    def expr(self, n):
        m = getattr(self, "e_" + type(n).__name__, None)
        if m is None:
            return f"/* TODO {type(n).__name__} */ undefined"
        return m(n)

    # -------------------------------------------------------------- expressions
    def e_Constant(self, n):
        v = n.value
        if v is None:
            return "null"
        if v is True:
            return "true"
        if v is False:
            return "false"
        if isinstance(v, str):
            return repr_js(v)
        if isinstance(v, bytes):
            return bytes_js(v)
        if isinstance(v, float):
            if v != v:
                return "NaN"
            if v in (float("inf"), float("-inf")):
                return "Infinity" if v > 0 else "-Infinity"
            return repr(v)
        if v is Ellipsis:
            return "undefined"
        if isinstance(v, int) and abs(v) > 2**53:
            return f"{v}"
        return repr(v)

    def e_Name(self, n):
        i = n.id
        if self.is_local(i):
            return ident(i)
        if self.module_aliases.get(i) in ("hashlib", "zipfile", "itertools", "collections", "base64", "pptx", "datetime"):
            return "$" + self.module_aliases[i]
        if i in self.camel_names:
            return camel(i)
        if i in ("self", "cls") and getattr(self, "in_method", False):
            return "this"
        return {"self": "self", "sys": "$sys", "str": "String", "int": "Number", "float": "Number", "bool": "Boolean",
                "list": "Array", "tuple": "Array", "dict": "Object", "bytes": "Uint8Array", "object": "Object",
                "None": "null"}.get(i, i)

    def e_Attribute(self, n):
        base = n.value
        if isinstance(base, ast.Name):
            mod = self.module_aliases.get(base.id)
            if mod == "io" and n.attr in ("BytesIO", "StringIO", "RawIOBase", "BufferedReader"):
                return "BytesIO"
            if mod == "pytest":
                return f"pytest.{n.attr}"
            if mod in ("hashlib", "zipfile", "itertools", "collections", "base64", "os", "io") and mod != "os" and mod != "io":
                return f"${mod}.{n.attr}"
            if mod == "sys" and n.attr == "maxsize":
                return "Number.MAX_SAFE_INTEGER"
            if mod == "os" and n.attr == "path":
                return "$path"
        if isinstance(base, ast.Attribute) and isinstance(base.value, ast.Name) and self.module_aliases.get(base.value.id) == "os":
            return f"$path.{n.attr}"
        b = self.expr(base)
        if n.attr == "__name__":
            return f"{b}.name"
        if n.attr == "__class__":
            return f"$type({b})"
        if n.attr == "value" and isinstance(base, ast.Name) and base.id in self.excinfo_names:
            return f"{b}.value"
        if n.attr == "args" and isinstance(base, ast.Attribute) and base.attr == "value":
            return f"[{b}.message]"
        if n.attr == "args" and isinstance(base, ast.Attribute) and base.attr in ("call_args",):
            return f"{b}.pyArgs"
        if n.attr == "args" and isinstance(base, ast.Subscript):
            return f"{b}.pyArgs"
        return f"{b}.{attr_name(n.attr)}"

    excinfo_names = set()

    def e_List(self, n):
        return "[" + ", ".join(self.expr(e) for e in n.elts) + "]"

    e_Tuple = e_List

    def e_Set(self, n):
        return "new Set([" + ", ".join(self.expr(e) for e in n.elts) + "])"

    def e_Starred(self, n):
        return "...$list(" + self.expr(n.value) + ")"

    def e_Dict(self, n):
        parts = []
        for k, v in zip(n.keys, n.values):
            if k is None:
                parts.append("..." + self.expr(v))
            elif isinstance(k, ast.Constant) and isinstance(k.value, str) and re.match(r"^[A-Za-z_$][\w$]*$", k.value):
                parts.append(f"{k.value}: {self.expr(v)}")
            elif isinstance(k, ast.Constant) and isinstance(k.value, (str, int)):
                parts.append(f"{repr_js(k.value) if isinstance(k.value, str) else repr_js(str(k.value))}: {self.expr(v)}")
            else:
                parts.append(f"[{self.expr(k)}]: {self.expr(v)}")
        return "$d({" + ", ".join(parts) + "})"

    def e_UnaryOp(self, n):
        if isinstance(n.op, ast.Not):
            return f"!$truthy({self.expr(n.operand)})"
        if isinstance(n.op, ast.USub):
            return f"-{self.expr(n.operand)}"
        if isinstance(n.op, ast.Invert):
            return f"~{self.expr(n.operand)}"
        return self.expr(n.operand)

    def e_BinOp(self, n):
        l, r = self.expr(n.left), self.expr(n.right)
        op = type(n.op)
        lit = lambda x: isinstance(x, ast.Constant) and isinstance(x.value, (int, float)) and not isinstance(x.value, bool)
        if op is ast.Add:
            if lit(n.left) or lit(n.right):
                return f"({l} + {r})"
            return f"$add({l}, {r})"
        if op is ast.Mod:
            return f"$mod({l}, {r})"
        if op is ast.Mult:
            if lit(n.left) and lit(n.right):
                return f"({l} * {r})"
            return f"$mul({l}, {r})"
        if op is ast.FloorDiv:
            return f"Math.floor({l} / {r})"
        if op is ast.Pow:
            return f"({l} ** {r})"
        if op is ast.BitOr:
            return f"$or({l}, {r})"
        sym = {ast.Sub: "-", ast.Div: "/", ast.BitAnd: "&", ast.BitXor: "^", ast.LShift: "<<", ast.RShift: ">>"}[op]
        return f"({l} {sym} {r})"

    def e_BoolOp(self, n):
        # python and/or return operands
        vals = [self.expr(v) for v in n.values]
        fn = "$and" if isinstance(n.op, ast.And) else "$orv"
        return f"{fn}(" + ", ".join(f"() => {v}" for v in vals) + ")"

    def e_Compare(self, n):
        out = []
        left = n.left
        for op, right in zip(n.ops, n.comparators):
            l, r = self.expr(left), self.expr(right)
            t = type(op)
            if t is ast.Eq:
                out.append(f"$eq({l}, {r})")
            elif t is ast.NotEq:
                out.append(f"!$eq({l}, {r})")
            elif t is ast.Is:
                out.append(f"$is({l}, {r})")
            elif t is ast.IsNot:
                out.append(f"!$is({l}, {r})")
            elif t is ast.In:
                out.append(f"$in({l}, {r})")
            elif t is ast.NotIn:
                out.append(f"!$in({l}, {r})")
            else:
                sym = {ast.Lt: "<", ast.LtE: "<=", ast.Gt: ">", ast.GtE: ">="}[t]
                out.append(f"$cmp({l}, {json_str(sym)}, {r})")
            left = right
        return " && ".join(out)

    def e_IfExp(self, n):
        return f"($truthy({self.expr(n.test)}) ? {self.expr(n.body)} : {self.expr(n.orelse)})"

    def e_Lambda(self, n):
        args = [a.arg for a in n.args.args]
        self.scopes.append(set(args))
        body = self.expr(n.body)
        self.scopes.pop()
        if isinstance(n.body, ast.Dict):
            body = f"({body})"
        return f"(({', '.join(ident(a) + ': any' for a in args)}) => {body})"

    def e_JoinedStr(self, n):
        s = ""
        for v in n.values:
            if isinstance(v, ast.Constant):
                s += v.value.replace("\\", "\\\\").replace("`", "\\`").replace("${", "\\${")
            else:
                inner = self.expr(v.value)
                if v.conversion == ord("r"):
                    inner = f"$repr({inner})"
                if v.format_spec is not None:
                    spec = "".join(x.value for x in v.format_spec.values if isinstance(x, ast.Constant))
                    s += "${$format(" + inner + ", " + json_str(spec) + ")}"
                else:
                    s += "${$str(" + inner + ")}"
        return "`" + s + "`"

    def e_Subscript(self, n):
        v = self.expr(n.value)
        sl = n.slice
        if isinstance(sl, ast.Slice):
            lo = self.expr(sl.lower) if sl.lower else "null"
            hi = self.expr(sl.upper) if sl.upper else "null"
            st = self.expr(sl.step) if sl.step else "null"
            return f"$slice({v}, {lo}, {hi}, {st})"
        return f"$get({v}, {self.expr(sl)})"

    def comp(self, n, kind):
        def build(gens, elt_fn):
            g = gens[0]
            tgt = self.target(g.target)
            names = target_names(g.target)
            self.scopes.append(set(names))
            it = f"$list({self.expr(g.iter)})"
            for cond in g.ifs:
                it += f".filter(({tgt}: any) => $truthy({self.expr(cond)}))"
            if len(gens) == 1:
                res = f"{it}.map(({tgt}: any) => {elt_fn()})"
            else:
                res = f"{it}.flatMap(({tgt}: any) => {build(gens[1:], elt_fn)})"
            self.scopes.pop()
            return res

        if kind == "dict":
            body = build(n.generators, lambda: f"[{self.expr(n.key)}, {self.expr(n.value)}]")
            return f"$d(Object.fromEntries({body}))"
        def elt():
            e = self.expr(n.elt)
            return f"({e})" if isinstance(n.elt, ast.Dict) else e
        body = build(n.generators, elt)
        return f"new Set({body})" if kind == "set" else body

    def e_ListComp(self, n):
        return self.comp(n, "list")

    e_GeneratorExp = e_ListComp

    def e_SetComp(self, n):
        return self.comp(n, "set")

    def e_DictComp(self, n):
        return self.comp(n, "dict")

    def target(self, t):
        if isinstance(t, (ast.Tuple, ast.List)):
            return "[" + ", ".join(self.target(e) for e in t.elts) + "]"
        if isinstance(t, ast.Name):
            return ident(t.id)
        if isinstance(t, ast.Starred):
            return "..." + self.target(t.value)
        return self.expr(t)

    # --------------------------------------------------------------------- calls
    def kw_obj(self, kws):
        parts = []
        for k in kws:
            if k.arg is None:
                parts.append("..." + self.expr(k.value))
            else:
                parts.append(f"{k.arg}: {self.expr(k.value)}")
        return "$kw({" + ", ".join(parts) + "})"

    def args_js(self, args, kws):
        parts = [self.expr(a) for a in args]
        if kws:
            parts.append(self.kw_obj(kws))
        return ", ".join(parts)

    def e_Call(self, n):
        f = n.func
        args = n.args
        kws = n.keywords

        if isinstance(f, ast.Name):
            b = f.id
            if not self.is_local(b) and b not in self.imports:
                if b in EXC:
                    cls = EXC[b] if EXC[b] != "null" else "Error"
                    return f"new {cls}({self.args_js(args, [])})"
                if b == "isinstance" and len(args) == 2 and isinstance(args[1], ast.Tuple):
                    return f"$isinstance({self.expr(args[0])}, [{', '.join(self.expr(e) for e in args[1].elts)}])"
                if b == "super":
                    return "super"
                if b in BUILTINS:
                    if b == "str" and args and isinstance(args[0], ast.Call) is False and False:
                        pass
                    return f"{BUILTINS[b]}({self.args_js(args, kws)})"

        # pytest helpers
        fname = self.expr(f)
        if fname == "pytest.param":
            return f"pytestParam({self.args_js(args, kws)})"
        if fname == "pytest.approx":
            return f"$approx({self.args_js(args, kws)})"
        if fname in ("BytesIO",):
            return f"new BytesIO({self.args_js(args, [])})"
        if fname == "path.join" or fname == "path.abspath" or fname == "path.dirname" or fname == "path.split" or fname == "path.splitext" or fname == "path.exists" or fname == "path.isdir":
            return f"{fname}({self.args_js(args, [])})"

        if isinstance(f, ast.Attribute):
            obj = self.expr(f.value)
            meth = f.attr
            # str methods on literals
            if meth == "join" and isinstance(f.value, ast.Constant):
                return f"$list({self.expr(args[0])}).join({obj})"
            if meth == "format" and isinstance(f.value, ast.Constant):
                return f"$pyformat({obj}, {self.kw_obj(kws) if kws else '[' + ', '.join(self.expr(a) for a in args) + ']'})"
            if meth in DISPATCH and obj != "$path" and (meth in DISPATCH_ALWAYS or not self.src_receiver(f.value)):
                return f"{DISPATCH[meth]}({obj}" + "".join(", " + p for p in ([self.expr(a) for a in args] + ([self.kw_obj(kws)] if kws else []))) + ")"
            m = attr_name(meth)
            if meth == "assert_has_calls" or meth.startswith("assert_"):
                return f"{obj}.{m}({self.args_js(args, kws)})"
            if kws or any(isinstance(a, ast.Starred) for a in args):
                return f"$m({obj}, {json_str(m)}" + "".join(", " + p for p in [self.expr(a) for a in args] + ([self.kw_obj(kws)] if kws else [])) + ")"
            return f"{obj}.{m}({self.args_js(args, [])})"

        # plain name call
        if isinstance(f, ast.Name) and f.id in self.local_funcs and not kws:
            return f"{fname}({self.args_js(args, [])})"
        if isinstance(f, ast.Name) and f.id.lstrip("_")[:1].isupper() or kws or not isinstance(f, ast.Name):
            return f"$call({fname}" + ("".join(", " + p for p in [self.expr(a) for a in args] + ([self.kw_obj(kws)] if kws else []))) + ")"
        return f"{fname}({self.args_js(args, [])})"

    def src_receiver(self, node):
        """True if `node` is clearly a pptx object (so `.append` etc. are its own methods)."""
        if isinstance(node, ast.Name):
            n = node.id
            return bool(re.search(r"(elm|element|_lst|txBody|sp|spTree|tbl|tr|tc|p|r|xml_rels|rels|chartSpace|plotArea|ser|cat|pic|grpSp|cxnSp|graphicFrame|bodyPr|pPr|rPr|parent|child|root|xfrm|tree|ln|fill|types_elm|relationships|relationship|pkg|package|part)$", n)) and not n.endswith(("_list", "_names", "_seq", "s_"))
        if isinstance(node, ast.Attribute):
            return node.attr.startswith("_") or node.attr in ("element", "_element", "spTree")
        return False

    # ---------------------------------------------------------------- statements
    def body(self, stmts, ind):
        lines = []
        for s in stmts:
            lines.extend(self.stmt(s, ind))
        return lines

    def stmt(self, s, ind):
        p = "  " * ind
        m = getattr(self, "s_" + type(s).__name__, None)
        if m is None:
            return [p + f"// TODO {type(s).__name__}: " + ast.unparse(s).replace("\n", " ")[:200]]
        return m(s, ind)

    def s_Expr(self, s, ind):
        if isinstance(s.value, ast.Constant) and isinstance(s.value.value, str):
            return []
        if isinstance(s.value, ast.Yield):
            return ["  " * ind + f"yield {self.expr(s.value.value) if s.value.value else ''};"]
        return ["  " * ind + self.expr(s.value) + ";"]

    def exc(self, node):
        if isinstance(node, ast.Tuple):
            return "[" + ", ".join(self.exc(e) for e in node.elts) + "]"
        name = ast.unparse(node).split(".")[-1]
        return EXC.get(name, name)

    def s_Assert(self, s, ind):
        p = "  " * ind
        t = s.test
        if "__doc__" in ast.unparse(t):
            return [p + "// python docstring assertion omitted (no __doc__ in JS)"]
        msg = f", {self.expr(s.msg)}" if s.msg else ""
        if isinstance(t, ast.Compare) and len(t.ops) == 1 and isinstance(t.ops[0], ast.Eq):
            return [p + f"$assertEqual({self.expr(t.left)}, {self.expr(t.comparators[0])}{msg});"]
        return [p + f"$assert({self.expr(t)}{msg});"]

    def s_Pass(self, s, ind):
        return []

    def s_Return(self, s, ind):
        return ["  " * ind + ("return " + self.expr(s.value) + ";" if s.value else "return;")]

    def s_Raise(self, s, ind):
        return ["  " * ind + f"throw {self.expr(s.exc) if s.exc else '_e'};"]

    def s_Delete(self, s, ind):
        out = []
        for t in s.targets:
            if isinstance(t, ast.Subscript):
                out.append("  " * ind + f"$del({self.expr(t.value)}, {self.expr(t.slice)});")
            elif isinstance(t, ast.Attribute):
                out.append("  " * ind + f"delete {self.expr(t)};")
            else:
                out.append("  " * ind + f"// del {ast.unparse(t)}")
        return out

    def assign_to(self, t, val, ind):
        p = "  " * ind
        if isinstance(t, ast.Subscript) and not isinstance(t.slice, ast.Slice):
            return [p + f"$set({self.expr(t.value)}, {self.expr(t.slice)}, {val});"]
        if isinstance(t, ast.Name):
            decl = "let " if self.declare(t.id) else ""
            return [p + f"{decl}{ident(t.id)} = {val};"]
        if isinstance(t, (ast.Tuple, ast.List)):
            names = target_names(t)
            all_names = all(isinstance(e, ast.Name) for e in t.elts)
            fresh = all_names and all([self.declare(nm) for nm in names])
            if not all_names:
                for nm in names:
                    self.declare(nm)
            return [p + f"{'let ' if fresh else ''}{self.target(t)} = $list({val});"]
        return [p + f"{self.expr(t)} = {val};"]

    def s_Assign(self, s, ind):
        val = self.expr(s.value)
        if len(s.targets) == 1:
            return self.assign_to(s.targets[0], val, ind)
        # `a = b = expr`: evaluate expr once, then assign it to each target (python order)
        self.tmp_count = getattr(self, "tmp_count", 0) + 1
        tmp = f"$t{self.tmp_count}"
        out = ["  " * ind + f"const {tmp} = {val};"]
        for t in s.targets:
            out += self.assign_to(t, tmp, ind)
        return out

    def s_AnnAssign(self, s, ind):
        return self.assign_to(s.target, self.expr(s.value), ind) if s.value else []

    def s_AugAssign(self, s, ind):
        t = s.target
        cur = self.expr(t) if not isinstance(t, ast.Subscript) else f"$get({self.expr(t.value)}, {self.expr(t.slice)})"
        op = type(s.op)
        r = self.expr(s.value)
        if op is ast.Add:
            val = f"$add({cur}, {r})"
        elif op is ast.Mult:
            val = f"$mul({cur}, {r})"
        elif op is ast.Mod:
            val = f"$mod({cur}, {r})"
        else:
            sym = {ast.Sub: "-", ast.Div: "/", ast.BitAnd: "&", ast.BitOr: "|", ast.FloorDiv: "/"}[op]
            val = f"({cur} {sym} {r})"
        return self.assign_to(t, val, ind)

    def s_If(self, s, ind):
        p = "  " * ind
        if "__name__" in ast.unparse(s.test) or ast.unparse(s.test) == "TYPE_CHECKING":
            return []
        out = [p + f"if ($truthy({self.expr(s.test)})) {{"] + self.body(s.body, ind + 1)
        if s.orelse:
            out.append(p + "} else {")
            out += self.body(s.orelse, ind + 1)
        out.append(p + "}")
        return out

    def s_For(self, s, ind):
        p = "  " * ind
        self.scopes.append(set(target_names(s.target)))
        out = [p + f"for (const {self.target(s.target)} of $list({self.expr(s.iter)})) {{"] + self.body(s.body, ind + 1) + [p + "}"]
        self.scopes.pop()
        return out

    def s_While(self, s, ind):
        p = "  " * ind
        return [p + f"while ($truthy({self.expr(s.test)})) {{"] + self.body(s.body, ind + 1) + [p + "}"]

    def s_Break(self, s, ind):
        return ["  " * ind + "break;"]

    def s_Continue(self, s, ind):
        return ["  " * ind + "continue;"]

    def s_Try(self, s, ind):
        p = "  " * ind
        out = [p + "try {"] + self.body(s.body, ind + 1)
        if s.handlers:
            h = s.handlers[0]
            out.append(p + f"}} catch ({h.name or '_e'}: any) {{")
            out += self.body(h.body, ind + 1)
        if s.finalbody:
            out.append(p + "} finally {")
            out += self.body(s.finalbody, ind + 1)
        out.append(p + "}")
        return out

    def s_With(self, s, ind):
        p = "  " * ind
        item = s.items[0]
        ce = item.context_expr
        if isinstance(ce, ast.Call):
            fn = ast.unparse(ce.func)
            if fn in ("pytest.raises",):
                exc = self.exc(ce.args[0])
                match = None
                for k in ce.keywords:
                    if k.arg == "match":
                        match = self.expr(k.value)
                inner = self.body(s.body, ind + 1)
                tail = "}" + (f", {match}" if match else "") + ");"
                if item.optional_vars is not None:
                    name = item.optional_vars.id
                    self.excinfo_names.add(name)
                    decl = "const " if self.declare(name) else ""
                    return [p + f"{decl}{ident(name)} = $raises({exc}, () => {{"] + inner + [p + tail]
                return [p + f"$raises({exc}, () => {{"] + inner + [p + tail]
            if fn in ("pytest.warns", "warnings.catch_warnings"):
                return [p + "{"] + self.body(s.body, ind + 1) + [p + "}"]
        # a context manager: `__enter__()` gives the bound value, `__exit__()` runs afterwards
        self.tmp_count = getattr(self, "tmp_count", 0) + 1
        cm = f"$cm{self.tmp_count}"
        out = [p + f"const {cm} = {self.expr(ce)};"]
        if item.optional_vars is not None and isinstance(item.optional_vars, ast.Name):
            name = item.optional_vars.id
            decl = "let " if self.declare(name) else ""
            out.append(p + f"{decl}{ident(name)} = $enter({cm});")
        else:
            out.append(p + f"$enter({cm});")
        return out + [p + "try {"] + self.body(s.body, ind + 1) + [p + "} finally {", p + f"  $exit({cm});", p + "}"]

    def s_Import(self, s, ind):
        for a in s.names:
            self.module_aliases[a.asname or a.name] = a.name
        return []

    def s_ImportFrom(self, s, ind):
        mod = ("." * s.level) + (s.module or "")
        if mod in ("typing", "__future__", "typing_extensions", "collections.abc"):
            return []
        for a in s.names:
            local = a.asname or a.name
            self.imports[local] = (mod, a.name)
            if mod == "unittest" and a.name == "mock":
                self.module_aliases[local] = "mock"
            if a.name[:1].islower() and (mod.startswith(CAMEL_MODULE_PREFIXES) or mod.startswith(".")):
                self.camel_names.add(local)
            elif a.name[:1].isupper():
                self.classes.add(local)
        return []

    def s_Global(self, s, ind):
        return []

    def s_ClassDef(self, s, ind):
        return self.helper_class(s, ind)

    def s_FunctionDef(self, s, ind):
        # nested function
        p = "  " * ind
        self.declare(s.name)
        self.local_funcs.add(s.name)
        return [p + f"const {s.name} = {self.func_expr(s, ind)};"]

    # ------------------------------------------------------- functions / tests
    def params(self, s, skip_self=True):
        a = s.args
        names = [x.arg for x in a.args]
        if skip_self and names and names[0] in ("self", "cls"):
            names = names[1:]
        defaults = [None] * (len(names) - len(a.defaults)) + list(a.defaults)
        out = []
        for nm, d in zip(names, defaults):
            out.append(f"{ident(nm)}: any" + (f" = {self.expr(d)}" if d is not None else ""))
        for k, d in zip(a.kwonlyargs, a.kw_defaults):
            out.append(f"{ident(k.arg)}: any" + (f" = {self.expr(d)}" if d is not None else ""))
        if a.vararg:
            out.append(f"...{ident(a.vararg.arg)}: any[]")
        if a.kwarg:
            out.append(f"{ident(a.kwarg.arg)}: any = {{}}")
        return out

    def func_body(self, s, ind, params):
        local = sorted(assigned_names(s.body) - set(params))
        self.scopes.append(set(params) | set(local))
        body = self.body(s.body, ind + 1)
        self.scopes.pop()
        if local:
            body = ["  " * (ind + 1) + "let " + ", ".join(f"{ident(x)}: any" for x in local) + ";"] + body
        return body

    def func_expr(self, s, ind, with_self=False):
        p = "  " * ind
        a = s.args
        params = [x.arg for x in a.args] + [k.arg for k in a.kwonlyargs] + ([a.vararg.arg] if a.vararg else []) + ([a.kwarg.arg] if a.kwarg else [])
        if params and params[0] in ("self", "cls") and not with_self:
            params = params[1:]
        plist = self.params(s, skip_self=not with_self)
        if with_self and plist and not plist[0].startswith("self"):
            pass
        body = self.func_body(s, ind, params)
        gen = has_yield(s)
        head = f"function* ({', '.join(plist)}): any {{" if gen else f"({', '.join(plist)}): any => {{"
        return head + "\n" + "\n".join(body) + "\n" + p + "}"

    def decorator_info(self, s):
        info = {"fixture": None, "parametrize": [], "skip": False, "xfail": False}
        for d in s.decorator_list:
            src = ast.unparse(d)
            if src in ("pytest.fixture", "fixture"):
                info["fixture"] = {}
            elif isinstance(d, ast.Call) and ast.unparse(d.func) in ("pytest.fixture", "fixture"):
                info["fixture"] = {k.arg: k.value for k in d.keywords}
            elif isinstance(d, ast.Call) and "parametrize" in ast.unparse(d.func):
                names = d.args[0]
                if isinstance(names, ast.Constant):
                    an = [x.strip() for x in names.value.split(",")]
                else:
                    an = [e.value for e in names.elts]
                ids = None
                for k in d.keywords:
                    if k.arg == "ids":
                        ids = k.value
                info["parametrize"].append((an, d.args[1], ids))
            elif "skip" in src:
                info["skip"] = True
            elif "xfail" in src:
                info["xfail"] = True
        return info

    def test_or_fixture(self, s, ind):
        """Return ('fixture'|'test'|'helper', js) for a method of a Describe class or module function."""
        info = self.decorator_info(s)
        a = s.args
        all_params = [x.arg for x in a.args]
        if all_params and all_params[0] == "self":
            all_params = all_params[1:]
        p = "  " * ind
        body = self.func_body(s, ind + 1, all_params)
        gen = has_yield(s)
        fn = ("function* " if gen else "") + f"(self: any{''.join(', ' + ident(x) + ': any' for x in all_params)})" + (": any {" if gen else ": any => {")
        fn_js = fn + "\n" + "\n".join(body) + "\n" + p + "  }"
        args_js = "[" + ", ".join(json_str(x) for x in all_params) + "]"
        if info["fixture"] is not None:
            fx = info["fixture"]
            extra = ""
            if "params" in fx:
                extra += f", params: {self.expr(fx['params'])}"
            if "ids" in fx:
                extra += f", ids: {self.expr(fx['ids'])}"
            if "scope" in fx:
                extra += f", scope: {self.expr(fx['scope'])}"
            if "autouse" in fx:
                extra += f", autouse: {self.expr(fx['autouse'])}"
            return "fixture", f"{p}{s.name}: {{ args: {args_js}{extra}, fn: {fn_js} }},"
        name = s.name
        is_test = name.startswith(("it_", "its_", "and_", "but_", "test_", "they_", "it ", "test")) or name.startswith("it")
        if is_test:
            extra = ""
            if info["parametrize"]:
                parts = []
                for an, vals, ids in reversed(info["parametrize"]):
                    parts.append("{ names: [" + ", ".join(json_str(x) for x in an) + "], values: " + self.expr(vals) + (f", ids: {self.expr(ids)}" if ids is not None else "") + " }")
                extra += ", parametrize: [" + ", ".join(parts) + "]"
            if info["skip"]:
                extra += ", skip: true"
            if info["xfail"]:
                extra += ", xfail: true"
            return "test", f"{p}{{ name: {json_str(name)}, args: {args_js}{extra}, fn: {fn_js} }},"
        # helper method on the test class
        plist = self.params(s)
        hp = [x.arg for x in a.args][1:] + [k.arg for k in a.kwonlyargs] + ([a.vararg.arg] if a.vararg else []) + ([a.kwarg.arg] if a.kwarg else [])
        hbody = self.func_body(s, ind + 1, hp)
        return "helper", f"{p}{attr_name(s.name)}: (self: any{''.join(', ' + x for x in plist)}): any => {{\n" + "\n".join(hbody) + "\n" + p + "  },"

    def describe_class(self, s, ind):
        fixtures, tests, helpers, attrs = [], [], [], []
        for st in s.body:
            if isinstance(st, ast.FunctionDef):
                kind, js = self.test_or_fixture(st, 2)
                {"fixture": fixtures, "test": tests, "helper": helpers}[kind].append(js)
            elif isinstance(st, ast.Assign) and isinstance(st.targets[0], ast.Name):
                attrs.append(f"    {st.targets[0].id}: {self.expr(st.value)},")
        base = None
        for b in s.bases:
            if isinstance(b, ast.Name) and b.id in self.test_classes:
                base = b.id
        # a def without `self` used while the class body runs (e.g. `@pytest.fixture(params=_cases())`)
        # is a plain function: emit it at module level, before the class object
        pre = []
        for st in s.body:
            if isinstance(st, ast.FunctionDef) and not st.args.args and not st.decorator_list:
                pre.append(f"const {st.name} = {self.func_expr(st, 0)};")
        out = pre + [f"const {s.name}: any = {{"]
        if base:
            out.append(f"  base: {base},")
        out.append("  fixtures: {")
        out += fixtures
        out.append("  },")
        out.append("  helpers: {")
        out += helpers
        out.append("  },")
        if attrs:
            out.append("  attrs: {")
            out += attrs
            out.append("  },")
        out.append("  tests: [")
        out += tests
        out.append("  ],")
        out.append("};")
        if s.name.startswith(("Describe", "Test")) and tests:
            out.append(f"pytestClass({json_str(s.name)}, {s.name});")
        return out

    def helper_class(self, s, ind):
        """A non-test class defined in a test module (e.g. a subclass under test)."""
        p = "  " * ind
        bases = [self.expr(b) for b in s.bases if ast.unparse(b) not in ("object",)]
        ext = f" extends {bases[0]}" if bases else ""
        self.classes.add(s.name)
        out = [p + f"{'export ' if ind == 0 else ''}class {s.name}{ext} {{"]
        defines = []
        for st in s.body:
            if isinstance(st, ast.FunctionDef):
                decos = [ast.unparse(d) for d in st.decorator_list]
                name = st.name
                if name == "__init__":
                    plist = self.params(st)
                    params = [x.arg for x in st.args.args][1:]
                    self.in_method = True
                    body = self.func_body(st, ind + 1, params)
                    self.in_method = False
                    if bases and not any("super" in l for l in body):
                        body = [p + "    super();"] + body
                    out.append(p + f"  constructor({', '.join(plist)}) {{")
                    out += [l.replace("super.__init__(", "super(") for l in body]
                    out.append(p + "  }")
                    continue
                static = "classmethod" in decos or "staticmethod" in decos
                prop = "property" in decos or "lazyproperty" in decos
                params = [x.arg for x in st.args.args] + [k.arg for k in st.args.kwonlyargs] + ([st.args.vararg.arg] if st.args.vararg else []) + ([st.args.kwarg.arg] if st.args.kwarg else [])
                if params and params[0] in ("self", "cls"):
                    params = params[1:]
                plist = self.params(st)
                self.in_method = True
                body = self.func_body(st, ind + 1, params)
                self.in_method = False
                jsname = attr_name(name)
                if name == "__eq__":
                    jsname = "equals"
                if name == "__iter__":
                    jsname = "[Symbol.iterator]"
                if name == "__len__":
                    jsname = "get length"
                kw = "static " if static else ""
                if prop:
                    out.append(p + f"  {kw}get {jsname}(): any {{")
                else:
                    out.append(p + f"  {kw}{'*' if has_yield(st) else ''}{jsname}({', '.join(plist)}): any {{")
                out += body
                out.append(p + "  }")
            elif isinstance(st, (ast.Assign, ast.AnnAssign)):
                tgt = st.targets[0] if isinstance(st, ast.Assign) else st.target
                if st.value is None:
                    continue
                v = st.value
                if isinstance(v, ast.Call) and ast.unparse(v.func).split(".")[-1] in ("ZeroOrOne", "ZeroOrMore", "OneAndOnlyOne", "OneOrMore", "OptionalAttribute", "RequiredAttribute", "ZeroOrOneChoice", "Choice"):
                    defines.append(f"{attr_name(tgt.id)}: {self.expr(v).replace('$call(', '$newm(')}")
                else:
                    out.append(p + f"  static {tgt.id}: any = {self.expr(v)};")
        if defines:
            out.append(p + "  static {")
            out.append(p + "    defineChildren(this, { " + ", ".join(defines) + " });")
            out.append(p + "  }")
        out.append(p + "}")
        return out


def has_yield(fn):
    for node in ast.walk(fn):
        if isinstance(node, (ast.Yield, ast.YieldFrom)):
            # ignore yields in nested functions
            return True
    return False


def target_names(t):
    if isinstance(t, ast.Name):
        return [t.id]
    if isinstance(t, (ast.Tuple, ast.List)):
        out = []
        for e in t.elts:
            out += target_names(e)
        return out
    if isinstance(t, ast.Starred):
        return target_names(t.value)
    return []


def assigned_names(stmts):
    """Names bound anywhere in a function body (python function scope)."""
    names = set()

    def visit(node):
        if isinstance(node, (ast.FunctionDef, ast.Lambda, ast.ClassDef, ast.ListComp, ast.DictComp, ast.SetComp, ast.GeneratorExp)):
            # nested `def`s are emitted as `const name = ...` (see s_FunctionDef)
            # nested classes are emitted as `class Name {...}` declarations
            return
        if isinstance(node, ast.Name) and isinstance(node.ctx, ast.Store):
            names.add(node.id)
        if isinstance(node, ast.For):
            for child in ast.iter_child_nodes(node):
                if child is not node.target:
                    visit(child)
            return
        if isinstance(node, ast.With):
            for it in node.items:
                if it.optional_vars is not None and isinstance(it.optional_vars, ast.Name):
                    names.add(it.optional_vars.id)
        for child in ast.iter_child_nodes(node):
            visit(child)

    for st in stmts:
        visit(st)
    return names


def import_lines(t, rel_src, rel_test):
    """ES imports for the names the python module imported."""
    by_mod = {}
    for local, (mod, orig) in t.imports.items():
        if mod.startswith("pptx"):
            path = mod[len("pptx"):].lstrip(".").replace(".", "/") or "index"
            src_root = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "src"))
            if os.path.isdir(os.path.join(src_root, path)):
                path += "/index"
            target = f"{rel_src}/{path}.js"
        elif mod.startswith(".") or mod.startswith("tests"):
            level = len(mod) - len(mod.lstrip("."))
            m = mod.lstrip(".")
            if m.startswith("tests."):
                m = m[len("tests."):]
            elif level:
                # relative to the test module's package (tests/<pkg>/...), level 1 = same package
                pkg = getattr(t, "pkg_parts", [])
                base = pkg[: len(pkg) - (level - 1)] if level > 1 else pkg
                m = ".".join(base + ([m] if m else []))
            if m.startswith("unitutil.mock") or m == "unitutil.mock":
                continue  # provided by pyshim
            if m in ("unitutil", "unitutil.__init__"):
                continue
            target = f"{rel_test}/{m.replace('.', '/')}.js"
        else:
            continue
        name = camel(orig) if local in t.camel_names else orig
        alias = camel(local) if local in t.camel_names else local
        by_mod.setdefault(target, []).append(name if name == alias else f"{name} as {alias}")
    lines = []
    snapshots = []
    for target, names in sorted(by_mod.items()):
        names = sorted(set(names))
        if target.startswith(rel_src + "/"):
            specs = []
            for nm in names:
                orig, _, alias = nm.partition(" as ")
                alias = alias or orig
                # initializer_mock replaces the class in its module bindings (JS can't patch a
                # constructor in place), so tests that use it need the live binding
                if alias in getattr(t, "initializer_mocked", set()):
                    specs.append(nm)
                    continue
                specs.append(f"{orig} as {alias}$imported")
                snapshots.append(f"const {alias} = {alias}$imported;")
            lines.append(f"import {{ {', '.join(specs)} }} from \"{target}\";")
        else:
            lines.append(f"import {{ {', '.join(names)} }} from \"{target}\";")
    if snapshots:
        lines.append("// python binds imported names at import time; patches made later don't change them")
        lines += snapshots
    return lines


def main():
    src, dest = sys.argv[1], sys.argv[2]
    tree = ast.parse(open(src, encoding="utf-8").read())
    t = T()
    # first pass: imports, names of classes and module-level functions
    for s in tree.body:
        if isinstance(s, (ast.Import, ast.ImportFrom)):
            t.stmt(s, 0)
        elif isinstance(s, ast.ClassDef):
            t.classes.add(s.name)
    lines = []
    module_fixtures, module_tests = [], []
    for s in tree.body:
        if isinstance(s, (ast.Import, ast.ImportFrom)):
            continue
        if isinstance(s, ast.ClassDef):
            is_test = s.name.startswith(("Describe", "Test")) or any(
                isinstance(st, ast.FunctionDef) and st.name.startswith(("it_", "test_")) for st in s.body
            )
            if is_test:
                t.test_classes[s.name] = s
                lines += t.describe_class(s, 0)
            else:
                lines += t.helper_class(s, 0)
        elif isinstance(s, ast.FunctionDef):
            info = t.decorator_info(s)
            if info["fixture"] is not None or s.name.startswith("test_") or s.name.startswith("it_"):
                kind, js = t.test_or_fixture(s, 1)
                (module_fixtures if kind == "fixture" else module_tests).append(js)
            else:
                t.local_funcs.add(s.name)
                lines.append(f"export function {s.name}({', '.join(t.params(s, skip_self=False))}): any {{")
                params = [x.arg for x in s.args.args] + ([s.args.vararg.arg] if s.args.vararg else []) + ([s.args.kwarg.arg] if s.args.kwarg else [])
                lines += t.func_body(s, 0, params)
                lines.append("}")
        else:
            lines += t.stmt(s, 0)
        lines.append("")
    if module_fixtures or module_tests:
        lines.append("pytestModule({")
        lines.append("  fixtures: {")
        lines += module_fixtures
        lines.append("  },")
        lines.append("  tests: [")
        lines += module_tests
        lines.append("  ],")
        lines.append("});")
    # module fixtures must be registered before classes run: move pytestModule to the top
    out_dir = os.path.dirname(os.path.abspath(dest))
    test_root = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "test"))
    src_root = os.path.abspath(os.path.join(test_root, "..", "src"))
    rel_test = os.path.relpath(test_root, out_dir).replace("\\", "/")
    rel_src = os.path.relpath(src_root, out_dir).replace("\\", "/")
    if not rel_test.startswith("."):
        rel_test = "./" + rel_test
    header = [
        f"// Machine-translated from python-pptx {src.replace(chr(92), '/').split('python-pptx/')[-1]} by",
        "// scripts/py2ts_tests.py, then reviewed. Python semantics come from test/pyshim.ts.",
        "/* eslint-disable */",
        "// @ts-nocheck",
        f'import * as $py from "{rel_test}/pyshim.js";',
        "__SHIM_IMPORTS__",
        f'import {{ AttributeError, FileNotFoundError, IndexError, KeyError, NotImplementedError, ValueError, ZeroDivisionError }} from "{rel_src}/errors.js";',
        f'import {{ InvalidXmlError, PackageNotFoundError, PythonPptxError }} from "{rel_src}/exc.js";',
        f'import {{ defineChildren }} from "{rel_src}/oxml/xmlchemy.js";',
    ]
    rel_py = src.replace("\\", "/").split("/tests/", 1)[-1]
    t.pkg_parts = rel_py.split("/")[:-1]
    t.initializer_mocked = set(re.findall(r"initializer_mock\(\s*request\s*,\s*([A-Za-z_]\w*)", open(src, encoding="utf-8").read()))
    header += import_lines(t, rel_src, rel_test)
    header.append("void $py;")
    header.append("")
    # put pytestModule (fixture registration) before the classes that use it
    if module_fixtures or module_tests:
        idx = lines.index("pytestModule({")
        block = lines[idx:]
        lines = block + [""] + lines[:idx]
    # import exactly the shim names the body uses
    body_text = "\n".join(lines)
    shim_names = (r"(?<![\w$.])(\$[A-Za-z_]\w*|ANY|BytesIO|call|classMock|clsAttrMock|createAutospec|functionMock|"
                  r"initializerMock|instanceMock|looseMock|MagicMock|methodMock|Mock|openMock|PropertyMock|"
                  r"propertyMock|pytest|pytestClass|pytestModule|pytestParam|varMock|patchObject)(?![\w$])")
    used = sorted(set(re.findall(shim_names, body_text)) - {"$py"})
    header = [f'import {{ {", ".join(used)} }} from "{rel_test}/pyshim.js";' if h == "__SHIM_IMPORTS__" else h for h in header]
    text = "\n".join(header + lines) + "\n"
    # recorded post-translation edits (test/fixups.json)
    import json
    fixups_path = os.path.join(test_root, "fixups.json")
    if os.path.exists(fixups_path):
        rel_dest = os.path.relpath(os.path.abspath(dest), os.path.dirname(test_root)).replace("\\", "/")
        for find, repl, *_ in json.load(open(fixups_path, encoding="utf-8")).get(rel_dest, []):
            if find == "@skip":
                # ["@skip", "<test name>", "<reason>"]: python-only behavior, see PARITY.md
                pat = '{ name: ' + json_str(repl) + ', '
                if pat not in text:
                    print(f"warning: no test {repl!r} to skip in {rel_dest}", file=sys.stderr)
                text = text.replace(pat, pat + f"skip: {json_str(_[0] if _ else 'N/A')}, ")
                continue
            if find not in text:
                print(f"warning: fixup not applied in {rel_dest}: {find[:60]!r}", file=sys.stderr)
            text = text.replace(find, repl)
    # fixups may use more shim names: recompute the shim import
    body_only = text.split("void $py;", 1)[1]
    used = sorted(set(re.findall(shim_names, body_only)) - {"$py"})
    text = re.sub(r'^import \{[^}]*\} from "[^"]*pyshim\.js";$', f'import {{ {", ".join(used)} }} from "{rel_test}/pyshim.js";', text, count=1, flags=re.M)
    open(dest, "w", encoding="utf-8").write(text)


if __name__ == "__main__":
    main()
