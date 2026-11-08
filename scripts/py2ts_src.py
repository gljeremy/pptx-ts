"""Draft a TypeScript module from a python-pptx source module.

This is a porting aid, not a compiler: it writes idiomatic-looking TypeScript (plain
operators, camelCase names, getters for properties, `lazy()` for lazyproperty, xmlchemy
declarations gathered into a `defineChildren()` static block, docstrings as JSDoc) so the
port starts from upstream's exact structure. Every draft is then reviewed and corrected by
hand before it goes into `src/`; constructs it can't map are left with a `/*TODO*/` marker.

Usage: python scripts/py2ts_src.py <python-pptx>/src/pptx/oxml/slide.py drafts/oxml/slide.ts
"""
import ast
import json
import os
import re
import sys

XMLCHEMY = {"ZeroOrOne", "ZeroOrMore", "OneAndOnlyOne", "OneOrMore", "ZeroOrOneChoice", "Choice",
            "OptionalAttribute", "RequiredAttribute"}
JS_RESERVED = {"new", "default", "function", "class", "delete", "in", "with", "package", "export",
               "import", "switch", "case", "typeof", "void", "arguments", "eval", "var", "this",
               "const", "let", "enum", "interface", "private", "public", "static", "yield", "super"}
TYPE_MAP = {"str": "string", "int": "number", "float": "number", "bool": "boolean", "bytes": "Uint8Array",
            "None": "null", "Any": "any", "object": "unknown", "Length": "number", "Emu": "number"}


def camel(name):
    m = re.match(r"^_*", name).group(0)
    rest = name[len(m):]
    if not rest or (name.startswith("__") and name.endswith("__")):
        return name
    if rest.isupper() or rest[0].isupper():
        return name
    return m + re.sub(r"_+([A-Za-z0-9])", lambda x: x.group(1).upper(), rest)


def ident(name):
    n = camel(name)
    return n + "_" if n in JS_RESERVED else n


def js_str(s):
    return json.dumps(s, ensure_ascii=False)


class Translator:
    def __init__(self, src_path):
        self.src_path = src_path
        self.src = open(src_path, encoding="utf-8").read()
        self.tree = ast.parse(self.src)
        self.out = []
        self.cls = None
        self.assign_counts = {}
        self.cls_name = None
        self.cls_statics = set()

    # ---- types -------------------------------------------------------------------------
    def ann(self, n):
        if n is None:
            return "any"
        if isinstance(n, ast.Constant):
            if n.value is None:
                return "null"
            if isinstance(n.value, str):
                try:
                    return self.ann(ast.parse(n.value, mode="eval").body)
                except SyntaxError:
                    return "any"
        if isinstance(n, ast.Name):
            return TYPE_MAP.get(n.id, n.id)
        if isinstance(n, ast.Attribute):
            return n.attr
        if isinstance(n, ast.BinOp) and isinstance(n.op, ast.BitOr):
            return f"{self.ann(n.left)} | {self.ann(n.right)}"
        if isinstance(n, ast.Subscript):
            base = self.ann(n.value)
            args = n.slice.elts if isinstance(n.slice, ast.Tuple) else [n.slice]
            a = [self.ann(x) for x in args]
            if base in ("list", "List", "Sequence", "Iterable", "Tuple", "tuple"):
                if base in ("Tuple", "tuple") and len(a) > 1 and "..." not in [ast.unparse(x) for x in args]:
                    return "[" + ", ".join(x for x in a if x != "...") + "]"
                return f"{a[0]}[]"
            if base in ("Iterator", "Generator"):
                return f"Generator<{a[0]}>"
            if base in ("dict", "Dict", "Mapping"):
                return f"Record<string, {a[1] if len(a) > 1 else 'any'}>"
            if base == "Optional":
                return f"{a[0]} | null"
            if base in ("Type",):
                return f"typeof {a[0]}"
            if base == "Callable":
                return f"(...args: any[]) => {a[-1]}"
            if base == "IO":
                return "FileSource"
            if base == "cast":
                return a[-1]
            return base
        return "any"

    # ---- expressions -------------------------------------------------------------------
    def expr(self, n):
        m = getattr(self, "e_" + type(n).__name__, None)
        if m is None:
            return f"/*TODO {type(n).__name__}*/"
        return m(n)

    def e_Constant(self, n):
        v = n.value
        if v is None:
            return "null"
        if v is True:
            return "true"
        if v is False:
            return "false"
        if v is Ellipsis:
            return "/*...*/"
        if isinstance(v, str):
            return js_str(v)
        if isinstance(v, bytes):
            return f"encodeUtf8({js_str(v.decode('latin-1'))})"
        return repr(v)

    def e_Name(self, n):
        if n.id == "self":
            return "this"
        if n.id == "cls":
            return "this"
        return ident(n.id)

    def e_Attribute(self, n):
        if n.attr == "__name__" and isinstance(n.value, ast.Call) and isinstance(n.value.func, ast.Name) and n.value.func.id == "type":
            return f"{self.expr(n.value.args[0])}.constructor.name"
        if isinstance(n.value, ast.Name) and n.value.id == "self" and self.cls_name and n.attr in self.cls_statics:
            return f"(this.constructor as typeof {self.cls_name}).{camel(n.attr)}"
        if isinstance(n.value, ast.Call) and isinstance(n.value.func, ast.Name) and n.value.func.id == "super":
            return f"super.{camel(n.attr)}"
        return f"{self.expr(n.value)}.{camel(n.attr)}"

    def e_Tuple(self, n):
        return "[" + ", ".join(self.expr(e) for e in n.elts) + "]"

    e_List = e_Tuple

    def e_Set(self, n):
        return "new Set([" + ", ".join(self.expr(e) for e in n.elts) + "])"

    def e_Starred(self, n):
        return "..." + self.expr(n.value)

    def e_Dict(self, n):
        items = []
        for k, v in zip(n.keys, n.values):
            if k is None:
                items.append("..." + self.expr(v))
            elif isinstance(k, ast.Constant) and isinstance(k.value, str):
                items.append(f"{js_str(k.value)}: {self.expr(v)}")
            else:
                items.append(f"[{self.expr(k)}]: {self.expr(v)}")
        return "{ " + ", ".join(items) + " }"

    def e_UnaryOp(self, n):
        o = self.expr(n.operand)
        if isinstance(n.op, ast.Not):
            simple = isinstance(n.operand, (ast.Name, ast.Attribute, ast.Subscript)) or (
                isinstance(n.operand, ast.Call) and not (isinstance(n.operand.func, ast.Name) and n.operand.func.id in ("isinstance", "len")))
            return f"!{o}" if simple else f"!({o})"
        if isinstance(n.op, ast.USub):
            return f"-{o}"
        if isinstance(n.op, ast.Invert):
            return f"~{o}"
        return o

    def e_BinOp(self, n):
        l, r = self.expr(n.left), self.expr(n.right)
        op = type(n.op)
        if op is ast.Mod and isinstance(n.left, ast.Constant) and isinstance(n.left.value, str):
            args = n.right.elts if isinstance(n.right, ast.Tuple) else [n.right]
            return self.percent_format(n.left.value, [self.expr(a) for a in args])
        if op is ast.FloorDiv:
            return f"Math.floor({l} / {r})"
        if op is ast.Pow:
            return f"{l} ** {r}"
        if op is ast.Mod:
            args = n.right.elts if isinstance(n.right, ast.Tuple) else [n.right]
            return f"pyPercent({l}, [{', '.join(self.expr(a) for a in args)}])"
        sym = {ast.Add: "+", ast.Sub: "-", ast.Mult: "*", ast.Div: "/", ast.BitAnd: "&", ast.BitOr: "|",
               ast.BitXor: "^", ast.LShift: "<<", ast.RShift: ">>"}[op]
        return f"({l} {sym} {r})"

    def percent_format(self, fmt, args):
        parts = re.split(r"(%[-#0 +]*\d*(?:\.\d+)?[sdrfxXi%])", fmt)
        s = ""
        i = 0
        for p in parts:
            if re.fullmatch(r"%[-#0 +]*\d*(?:\.\d+)?[sdrfxXi%]", p or ""):
                if p == "%%":
                    s += "%"
                    continue
                a = args[i] if i < len(args) else "/*TODO arg*/"
                i += 1
                s += "${" + (a if p[-1] in "sdi" else f"/*TODO {p}*/{a}") + "}"
            else:
                s += p.replace("\\", "\\\\").replace("`", "\\`").replace("${", "\\${")
        return "`" + s + "`"

    def e_BoolOp(self, n):
        sym = " && " if isinstance(n.op, ast.And) else " || "
        return "(" + sym.join(self.expr(v) for v in n.values) + ")"

    def e_Compare(self, n):
        out = []
        left = n.left
        for op, right in zip(n.ops, n.comparators):
            l, r = self.expr(left), self.expr(right)
            t = type(op)
            if t is ast.Eq:
                out.append(f"{l} === {r}")
            elif t is ast.NotEq:
                out.append(f"{l} !== {r}")
            elif t is ast.Is:
                out.append(f"{l} === {r}")
            elif t is ast.IsNot:
                out.append(f"{l} !== {r}")
            elif t is ast.In:
                out.append(f"{r}.includes({l})")
            elif t is ast.NotIn:
                out.append(f"!{r}.includes({l})")
            else:
                sym = {ast.Lt: "<", ast.LtE: "<=", ast.Gt: ">", ast.GtE: ">="}[t]
                out.append(f"{l} {sym} {r}")
            left = right
        return " && ".join(out)

    def e_IfExp(self, n):
        return f"({self.expr(n.test)} ? {self.expr(n.body)} : {self.expr(n.orelse)})"

    def e_Lambda(self, n):
        args = ", ".join(ident(a.arg) for a in n.args.args)
        return f"({args}) => {self.expr(n.body)}"

    def e_JoinedStr(self, n):
        s = ""
        for v in n.values:
            if isinstance(v, ast.Constant):
                s += v.value.replace("\\", "\\\\").replace("`", "\\`").replace("${", "\\${")
            else:
                inner = self.expr(v.value)
                if v.format_spec is not None:
                    inner = f"/*TODO fmt*/{inner}"
                s += "${" + inner + "}"
        return "`" + s + "`"

    def e_Subscript(self, n):
        v = self.expr(n.value)
        sl = n.slice
        if isinstance(sl, ast.Slice):
            lo = self.expr(sl.lower) if sl.lower else "0"
            if sl.step:
                return f"/*TODO step slice*/{v}"
            return f"{v}.slice({lo}{', ' + self.expr(sl.upper) if sl.upper else ''})"
        if isinstance(sl, ast.UnaryOp) and isinstance(sl.op, ast.USub):
            return f"{v}.at({self.expr(sl)})"
        return f"{v}[{self.expr(sl)}]"

    def comp(self, n):
        if len(n.generators) == 2 and not isinstance(n, ast.DictComp):
            g1, g2 = n.generators
            t1, t2 = self.expr(g1.target), self.expr(g2.target)
            inner = f"[...{self.expr(g2.iter)}]"
            for cond in g2.ifs:
                inner += f".filter(({t2}) => {self.expr(cond)})"
            inner += f".map(({t2}) => {self.expr(n.elt)})"
            outer = f"[...{self.expr(g1.iter)}]"
            for cond in g1.ifs:
                outer += f".filter(({t1}) => {self.expr(cond)})"
            return f"{outer}.flatMap(({t1}) => {inner})"
        gen = n.generators[0]
        tgt = self.expr(gen.target)
        it = self.expr(gen.iter)
        s = f"[...{it}]" if not isinstance(gen.iter, (ast.List, ast.Tuple)) else it
        for cond in gen.ifs:
            s += f".filter(({tgt}) => {self.expr(cond)})"
        elt = n.elt if not isinstance(n, ast.DictComp) else None
        if elt is not None:
            e = self.expr(elt)
            if e != tgt:
                s += f".map(({tgt}) => {e})"
        else:
            s = f"Object.fromEntries({s}.map(({tgt}) => [{self.expr(n.key)}, {self.expr(n.value)}]))"
        if len(n.generators) > 1:
            s = "/*TODO nested comp*/" + s
        return s

    e_ListComp = comp
    e_GeneratorExp = comp
    e_DictComp = comp

    def e_SetComp(self, n):
        return f"new Set({self.comp(n)})"

    def e_Yield(self, n):
        return f"yield {self.expr(n.value)}" if n.value is not None else "yield"

    def e_YieldFrom(self, n):
        return f"yield* {self.expr(n.value)}"

    def e_NamedExpr(self, n):
        return f"({self.expr(n.target)} = {self.expr(n.value)})"

    def kwargs(self, kws):
        return "{ " + ", ".join(f"{camel(k.arg)}: {self.expr(k.value)}" if k.arg else "..." + self.expr(k.value) for k in kws) + " }"

    def e_Call(self, n):
        f = n.func
        args = [self.expr(a) for a in n.args]
        kws = n.keywords
        if isinstance(f, ast.Name):
            name = f.id
            if name == "len" and len(args) == 1:
                return f"{args[0]}.length"
            if name == "isinstance":
                cls = n.args[1]
                prim = {"int": "number", "float": "number", "Number": "number", "str": "string", "bool": "boolean"}
                cname = cls.attr if isinstance(cls, ast.Attribute) else cls.id if isinstance(cls, ast.Name) else None
                if cname in prim:
                    return f'typeof {args[0]} === "{prim[cname]}"'
                if isinstance(cls, ast.Tuple) and all(isinstance(c, (ast.Name, ast.Attribute)) for c in cls.elts):
                    parts = []
                    for c in cls.elts:
                        cn = c.attr if isinstance(c, ast.Attribute) else c.id
                        parts.append(f'typeof {args[0]} === "{prim[cn]}"' if cn in prim else f"{args[0]} instanceof {self.expr(c)}")
                    return "(" + " || ".join(parts) + ")"
                if isinstance(cls, ast.Tuple):
                    return "(" + " || ".join(f"{args[0]} instanceof {self.expr(c)}" for c in cls.elts) + ")"
                return f"{args[0]} instanceof {args[1]}"
            if name == "cast":
                return args[1]
            if name == "str":
                return f"String({args[0]})" if args else '""'
            if name == "int":
                return f"pyInt({args[0]})"
            if name == "float":
                return f"pyFloat({args[0]})"
            if name == "bool":
                return f"pyBool({args[0]})"
            if name == "list":
                return f"[...{args[0]}]" if args else "[]"
            if name == "tuple":
                return f"[...{args[0]}]"
            if name == "dict" and not args:
                return self.kwargs(kws) if kws else "{}"
            if name == "dict" and len(args) == 1 and not kws:
                return f"Object.fromEntries({args[0]})"
            if name == "set":
                return f"new Set({args[0] if args else ''})"
            if name == "max":
                return f"Math.max({', '.join(args) if len(args) > 1 else '...' + args[0]})"
            if name == "min":
                return f"Math.min({', '.join(args) if len(args) > 1 else '...' + args[0]})"
            if name == "abs":
                return f"Math.abs({args[0]})"
            if name == "round":
                return f"pyRound({', '.join(args)})"
            if name == "sorted":
                kw = {k.arg: k.value for k in kws}
                key = self.expr(kw["key"]) if "key" in kw else None
                cmp = f"pyCmp(({key})(a), ({key})(b))" if key else "pyCmp(a, b)"
                if "reverse" in kw:
                    cmp = f"-{cmp}"
                return f"[...{args[0]}].sort((a, b) => {cmp})"
            if name == "next":
                return f"/*TODO next*/{args[0]}"
            if name == "iter":
                return f"{args[0]}[Symbol.iterator]()"
            if name == "enumerate":
                return f"[...{args[0]}].entries()"
            if name == "range":
                return f"range({', '.join(args)})"
            if name == "super":
                return "super"
            if name == "setattr":
                return f"(({args[0]}) as any)[{args[1]}] = {args[2]}"
            if name == "cls":
                return f"new this({', '.join(args + ([self.kwargs(kws)] if kws else []))})"
            if name == "getattr":
                return f"({args[0]} as any)[{args[1]}]" if len(args) == 2 else f"(({args[0]} as any)[{args[1]}] ?? {args[2]})"
            if name == "hasattr":
                return f"{args[1]} in {args[0]}"
            if name == "print":
                return f"console.log({', '.join(args)})"
            if name == "repr":
                return f"pyRepr({args[0]})"
            if name.lstrip("_")[:1].isupper() and not name.lstrip("_").isupper() and name not in ("Emu", "Inches", "Pt", "Cm", "Mm", "Centipoints", "PartFactory", "OxmlElement", "Presentation"):
                # a class: construct it
                a = args + ([self.kwargs(kws)] if kws else [])
                return f"new {name}({', '.join(a)})"
        callee = self.expr(f)
        if isinstance(f, ast.Attribute):
            meth = f.attr
            recv = self.expr(f.value)
            if meth == "join" and isinstance(f.value, ast.Constant):
                return f"{args[0]}.join({recv})"
            if meth == "format" and isinstance(f.value, ast.Constant):
                return f"/*TODO format*/{recv}"
            if meth == "isdigit":
                return r"/^\d+$/.test(" + recv + ")"
            if meth == "append" and "/oxml/" not in self.src_path.replace("\\", "/"):
                return f"{recv}.push({args[0]})"
            if meth == "startswith":
                return f"{recv}.startsWith({args[0]})"
            if meth == "endswith":
                return f"{recv}.endsWith({args[0]})"
            if meth == "strip":
                return f"{recv}.trim()"
            if meth == "lower":
                return f"{recv}.toLowerCase()"
            if meth == "upper":
                return f"{recv}.toUpperCase()"
            if meth == "items":
                return f"Object.entries({recv})"
            if meth == "values" and not args:
                return f"Object.values({recv})"
            if meth == "keys" and not args:
                return f"Object.keys({recv})"
            if isinstance(f.value, ast.Call) and isinstance(f.value.func, ast.Name) and f.value.func.id == "super" and meth == "__init__":
                return f"super({', '.join(args)})"
        a = args + ([self.kwargs(kws)] if kws else [])
        return f"{callee}({', '.join(a)})"

    # ---- statements --------------------------------------------------------------------
    def block(self, stmts, ind):
        lines = []
        for s in stmts:
            lines.extend(self.stmt(s, ind))
        return lines

    def stmt(self, s, ind):
        m = getattr(self, "s_" + type(s).__name__, None)
        pad = "  " * ind
        if m is None:
            return [f"{pad}/*TODO {type(s).__name__}*/"]
        return m(s, ind)

    def s_Expr(self, s, ind):
        if isinstance(s.value, ast.Constant) and isinstance(s.value.value, str):
            return []  # docstring
        return ["  " * ind + self.expr(s.value) + ";"]

    def s_Pass(self, s, ind):
        return []

    def s_Return(self, s, ind):
        return ["  " * ind + ("return " + self.expr(s.value) if s.value else "return") + ";"]

    def s_Raise(self, s, ind):
        if s.exc is None:
            return ["  " * ind + "throw e;"]
        e = self.expr(s.exc)
        if not e.startswith("new "):
            e = f"new {e}()" if re.fullmatch(r"\w+", e) else e
        return ["  " * ind + f"throw {e};"]

    def s_Assign(self, s, ind):
        pad = "  " * ind
        val = self.expr(s.value)
        out = []
        for t in s.targets:
            if isinstance(t, ast.Name) and t.id not in self.locals:
                self.locals.add(t.id)
                kw = "let" if self.assign_counts.get(t.id, 0) > 1 else "const"
                out.append(f"{pad}{kw} {ident(t.id)} = {val};")
            elif isinstance(t, ast.Tuple):
                names = ["" if isinstance(e, ast.Name) and e.id == "_" else self.expr(e) for e in t.elts]
                fresh = all(isinstance(e, ast.Name) and (e.id not in self.locals or e.id == "_") for e in t.elts)
                for e in t.elts:
                    if isinstance(e, ast.Name):
                        self.locals.add(e.id)
                out.append(f"{pad}{'const ' if fresh else ''}[{', '.join(names)}] = {val};")
            else:
                out.append(f"{pad}{self.expr(t)} = {val};")
        return out

    def s_AnnAssign(self, s, ind):
        if s.value is None:
            return []
        if isinstance(s.target, ast.Name) and s.target.id not in self.locals:
            self.locals.add(s.target.id)
            kw = "let" if self.assign_counts.get(s.target.id, 0) > 1 else "const"
            return ["  " * ind + f"{kw} {ident(s.target.id)}: {self.ann(s.annotation)} = {self.expr(s.value)};"]
        return ["  " * ind + f"{self.expr(s.target)} = {self.expr(s.value)};"]

    def s_AugAssign(self, s, ind):
        sym = {ast.Add: "+=", ast.Sub: "-=", ast.Mult: "*=", ast.Div: "/=", ast.BitOr: "|=", ast.BitAnd: "&="}.get(type(s.op), "/*TODO op*/=")
        return ["  " * ind + f"{self.expr(s.target)} {sym} {self.expr(s.value)};"]

    def s_If(self, s, ind):
        pad = "  " * ind
        out = [f"{pad}if ({self.expr(s.test)}) {{"]
        out += self.block(s.body, ind + 1)
        orelse = s.orelse
        while len(orelse) == 1 and isinstance(orelse[0], ast.If):
            out.append(f"{pad}}} else if ({self.expr(orelse[0].test)}) {{")
            out += self.block(orelse[0].body, ind + 1)
            orelse = orelse[0].orelse
        if orelse:
            out.append(f"{pad}}} else {{")
            out += self.block(orelse, ind + 1)
        out.append(f"{pad}}}")
        return out

    def s_For(self, s, ind):
        pad = "  " * ind
        tgt = self.expr(s.target)
        it = s.iter
        if isinstance(it, ast.Call) and isinstance(it.func, ast.Name) and it.func.id == "range":
            a = [self.expr(x) for x in it.args]
            lo, hi = ("0", a[0]) if len(a) == 1 else (a[0], a[1])
            head = f"for (let {tgt} = {lo}; {tgt} < {hi}; {tgt}++) {{"
        else:
            head = f"for (const {tgt} of {self.expr(it)}) {{"
        return [pad + head] + self.block(s.body, ind + 1) + [pad + "}"]

    def s_While(self, s, ind):
        pad = "  " * ind
        return [f"{pad}while ({self.expr(s.test)}) {{"] + self.block(s.body, ind + 1) + [pad + "}"]

    def s_Break(self, s, ind):
        return ["  " * ind + "break;"]

    def s_Continue(self, s, ind):
        return ["  " * ind + "continue;"]

    def s_Try(self, s, ind):
        pad = "  " * ind
        hoisted = []
        for st in s.body + [x for h in s.handlers for x in h.body]:
            for n in ast.walk(st):
                if isinstance(n, ast.Assign):
                    for t in n.targets:
                        for x in (t.elts if isinstance(t, ast.Tuple) else [t]):
                            if isinstance(x, ast.Name) and x.id not in self.locals and x.id not in hoisted:
                                hoisted.append(x.id)
        out = []
        if hoisted:
            out.append(pad + "let " + ", ".join(ident(h) for h in hoisted) + ";")
            self.locals.update(hoisted)
        out += [pad + "try {"] + self.block(s.body, ind + 1)
        for h in s.handlers:
            out.append(pad + "} catch (e) {")
            if h.type is not None:
                out.append(pad + f"  if (!(e instanceof {self.expr(h.type)})) throw e;")
            out += self.block(h.body, ind + 1)
        if s.finalbody:
            out.append(pad + "} finally {")
            out += self.block(s.finalbody, ind + 1)
        out.append(pad + "}")
        return out

    def s_With(self, s, ind):
        pad = "  " * ind
        return [pad + "/*TODO with " + ", ".join(ast.unparse(i) for i in s.items) + "*/"] + self.block(s.body, ind)

    def s_Delete(self, s, ind):
        return ["  " * ind + f"delete {self.expr(t)};" for t in s.targets]

    def s_Assert(self, s, ind):
        return ["  " * ind + f"if (!({self.expr(s.test)})) throw new Error(\"AssertionError\");"]

    def s_Import(self, s, ind):
        return []

    def s_ImportFrom(self, s, ind):
        return []

    def s_FunctionDef(self, s, ind):
        # nested function
        pad = "  " * ind
        params = self.params(s, skip_self=False)
        saved = self.locals
        self.locals = set(a.arg for a in s.args.args)
        body = self.block(s.body, ind + 1)
        self.locals = saved
        if has_yield(s):
            return [f"{pad}const {ident(s.name)} = function* (this: any, {params}) {{"] + body + [pad + "}.bind(this);"]
        return [f"{pad}const {ident(s.name)} = ({params}) => {{"] + body + [pad + "};"]

    # ---- definitions -------------------------------------------------------------------
    def params(self, s, skip_self=True):
        args = s.args.args
        if skip_self and args and args[0].arg in ("self", "cls"):
            args = args[1:]
        defaults = [None] * (len(args) - len(s.args.defaults)) + list(s.args.defaults)
        out = []
        for a, d in zip(args, defaults):
            p = f"{ident(a.arg)}: {self.ann(a.annotation)}"
            if d is not None:
                p = f"{ident(a.arg)}: {self.ann(a.annotation)} = {self.expr(d)}"
            out.append(p)
        for a, d in zip(s.args.kwonlyargs, s.args.kw_defaults):
            out.append(f"{ident(a.arg)}: {self.ann(a.annotation)}" + (f" = {self.expr(d)}" if d is not None else ""))
        if s.args.vararg:
            out.append(f"...{ident(s.args.vararg.arg)}: any[]")
        return ", ".join(out)

    def jsdoc(self, node, ind):
        doc = ast.get_docstring(node)
        if not doc:
            return []
        pad = "  " * ind
        lines = doc.strip().splitlines()
        if len(lines) == 1:
            return [f"{pad}/** {lines[0]} */"]
        return [f"{pad}/**"] + [f"{pad} * {l}".rstrip() for l in lines] + [f"{pad} */"]

    def method(self, s, ind):
        pad = "  " * ind
        decos = [ast.unparse(d) for d in s.decorator_list]
        name = camel(s.name)
        self.locals = set(a.arg for a in s.args.args + s.args.kwonlyargs)
        self.assign_counts = assign_counts(s)
        if s.args.vararg:
            self.locals.add(s.args.vararg.arg)
        multi = [n for n, c in self.assign_counts.items() if c > 1 and n not in self.locals and n != "_"]
        self.locals.update(multi)
        ret = self.ann(s.returns) if s.returns is not None else None
        body = self.block(s.body, ind + 2 if "lazyproperty" in decos else ind + 1)
        if multi:
            body = ["  " * (ind + 2 if "lazyproperty" in decos else ind + 1) + "let " + ", ".join(ident(m) for m in multi) + ";"] + body
        doc = self.jsdoc(s, ind)
        star = "*" if has_yield(s) else ""
        rt = f": {ret}" if ret and ret != "null" and s.name != "__init__" else (": void" if ret == "null" else "")
        if s.name == "__init__":
            if not self.cls_has_base:
                body = [l for l in body if l.strip() != "super();"]
            return doc + [f"{pad}constructor({self.params(s)}) {{"] + body + [pad + "}"]
        if s.name == "__eq__":
            return doc + [f"{pad}equals({self.params(s)}): boolean {{"] + body + [pad + "}"]
        if s.name in ("__ne__", "__hash__", "__repr__"):
            return []
        if s.name == "__iter__":
            return doc + [f"{pad}*[Symbol.iterator](){rt.replace('Iterator<', 'Generator<')} {{"] + body + [pad + "}"]
        if s.name == "__len__":
            return doc + [f"{pad}get length(): number {{"] + body + [pad + "}"]
        if s.name == "__getitem__":
            return doc + [f"{pad}at({self.params(s)}){rt} {{"] + body + [pad + "}"]
        if s.name == "__setitem__":
            return doc + [f"{pad}set({self.params(s)}): void {{"] + body + [pad + "}"]
        if s.name == "__contains__":
            return doc + [f"{pad}has({self.params(s)}): boolean {{"] + body + [pad + "}"]
        if "property" in decos:
            return doc + [f"{pad}get {name}(){rt} {{"] + body + [pad + "}"]
        if any(d.endswith(".setter") for d in decos):
            return [f"{pad}set {name}({self.params(s)}) {{"] + body + [pad + "}"]
        if "lazyproperty" in decos:
            return doc + [f"{pad}get {name}(){rt} {{", f'{pad}  return lazy(this, "{name}", () => {{'] + body + [pad + "  });", pad + "}"]
        static = "static " if ("classmethod" in decos or "staticmethod" in decos) else ""
        return doc + [f"{pad}{static}{star}{name}({self.params(s)}){rt} {{"] + body + [pad + "}"]

    def member_declares(self, pyname, call, anns):
        """{name: type} of the members xmlchemy generates for `pyname = <call>`."""
        kind = call.func.id
        p = camel(pyname)
        cap = lambda n: n[0].upper() + n[1:]
        t = anns.get(pyname, "any")
        el = re.sub(r"\s*\|\s*null$", "", t) if t != "any" else "any"
        if el.endswith("[]"):
            el = el[:-2]
        out = {}
        if kind in ("OptionalAttribute", "RequiredAttribute"):
            out[p] = t
        elif kind == "OneAndOnlyOne":
            out[p] = el
        elif kind == "ZeroOrOne":
            out[p] = f"{el} | null" if el != "any" else "any"
            out[f"getOrAdd{cap(p)}"] = f"() => {el}"
            out[f"_add{cap(p)}"] = f"(attrs?: Record<string, any>) => {el}"
            out[f"_insert{cap(p)}"] = f"(child: {el}) => {el}"
            out[f"_new{cap(p)}"] = f"() => {el}"
            out[f"_remove{cap(p)}"] = "() => void"
        elif kind in ("ZeroOrMore", "OneOrMore"):
            out[f"{p}Lst"] = f"{el}[]"
            out[f"_add{cap(p)}"] = f"(attrs?: Record<string, any>) => {el}"
            out[f"_insert{cap(p)}"] = f"(child: {el}) => {el}"
            out[f"_new{cap(p)}"] = f"() => {el}"
            if kind == "OneOrMore":
                out[f"add{cap(p)}"] = f"() => {el}"
        elif kind == "ZeroOrOneChoice":
            out[p] = "any"
            out[f"_remove{cap(p)}"] = "() => void"
            choices = call.args[0].elts if call.args and isinstance(call.args[0], (ast.Tuple, ast.List)) else []
            for c in choices:
                tag = c.args[0].value
                local = tag.split(":")[-1]
                out[local] = "any"
                out[f"_new{cap(local)}"] = "() => any"
                out[f"_insert{cap(local)}"] = "(child: any) => any"
                out[f"_add{cap(local)}"] = "(attrs?: Record<string, any>) => any"
                out[f"getOrChangeTo{cap(local)}"] = "() => any"
        return out

    def xmlchemy_call(self, call, tag_seq_name):
        """`ZeroOrOne("p:x", successors=_tag_seq[1:])` -> `new ZeroOrOne("p:x", _tagSeq.slice(1))`."""
        fname = call.func.id
        args = [self.expr(a) for a in call.args]
        kw = {k.arg: self.expr(k.value) for k in call.keywords}
        if fname in ("ZeroOrOne", "ZeroOrMore", "OneOrMore", "ZeroOrOneChoice"):
            if "successors" in kw:
                args.append(kw["successors"])
        elif fname == "OptionalAttribute":
            if "default" in kw:
                args.append(kw["default"])
        return f"new {fname}({', '.join(args)})"

    def class_def(self, s):
        self.cls_name = s.name
        self.cls_has_base = any(ast.unparse(b) not in ("object",) for b in s.bases)
        self.cls_statics = {
            st.name for st in s.body
            if isinstance(st, ast.FunctionDef) and any(ast.unparse(d) in ("classmethod", "staticmethod") for d in st.decorator_list)
        } | {
            st.targets[0].id for st in s.body
            if isinstance(st, ast.Assign) and len(st.targets) == 1 and isinstance(st.targets[0], ast.Name)
            and not (isinstance(st.value, ast.Call) and isinstance(st.value.func, ast.Name) and st.value.func.id in XMLCHEMY)
        }
        out = self.jsdoc(s, 0)
        bases = [ast.unparse(b) for b in s.bases if ast.unparse(b) not in ("object",)]
        ext = f" extends {bases[0]}" if bases else ""
        out.append(f"export class {s.name}{ext} {{")
        members = []
        declared = {}
        static_lines = []
        fields = []
        anns = {}
        for st in s.body:
            if isinstance(st, ast.AnnAssign) and isinstance(st.target, ast.Name):
                anns[st.target.id] = self.ann(st.annotation)
        for st in s.body:
            value = None
            tname = None
            if isinstance(st, ast.Assign) and len(st.targets) == 1 and isinstance(st.targets[0], ast.Name):
                tname, value = st.targets[0].id, st.value
            elif isinstance(st, ast.AnnAssign) and isinstance(st.target, ast.Name):
                tname, value = st.target.id, st.value
            if tname is None:
                continue
            if tname == "_tag_seq":
                static_lines.append(f"    const _tagSeq = {self.expr(value)} as const;")
                continue
            if isinstance(value, ast.Call) and isinstance(value.func, ast.Name) and value.func.id in XMLCHEMY:
                members.append((camel(tname), self.xmlchemy_call(value, None)))
                for k, v in self.member_declares(tname, value, anns).items():
                    declared.setdefault(k, v)
                continue
            if value is None:
                declared[camel(tname)] = anns[tname]
                continue
            self.locals = set()
            if isinstance(st, ast.AnnAssign):
                fields.append(f"  {camel(tname)}: {anns[tname]} = {self.expr(value)};")
            else:
                fields.append(f"  static {camel(tname)} = {self.expr(value)};")
        # attributes assigned on self in methods -> declared fields
        for st in s.body:
            if isinstance(st, ast.FunctionDef):
                for node in ast.walk(st):
                    targets = node.targets if isinstance(node, ast.Assign) else [node.target] if isinstance(node, ast.AnnAssign) else []
                    for t in targets:
                        if isinstance(t, ast.Attribute) and isinstance(t.value, ast.Name) and t.value.id == "self":
                            nm = camel(t.attr)
                            if nm not in declared:
                                declared[nm] = self.ann(node.annotation) if isinstance(node, ast.AnnAssign) else "any"
        # methods defined in the class body override generated members
        defined = {camel(st.name) for st in s.body if isinstance(st, ast.FunctionDef)}
        declared = {k: v for k, v in declared.items() if k not in defined}
        # annotations given upstream win over generated ones
        for k, v in anns.items():
            if camel(k) in declared:
                declared[camel(k)] = v
        out += [f"  declare {k}: {v};" for k, v in declared.items()] + fields
        if members:
            out.append("  static {")
            out += static_lines
            out.append("    defineChildren(this, {")
            for name, call in members:
                out.append(f"      {name}: {call},")
            out.append("    });")
            out.append("  }")
        for st in s.body:
            if isinstance(st, ast.FunctionDef):
                out.append("")
                out += self.method(st, 1)
        out.append("}")
        return out

    def run(self):
        doc = ast.get_docstring(self.tree)
        rel = os.path.relpath(self.src_path, os.path.join(os.path.dirname(self.src_path), "..")).replace("\\", "/")
        if doc:
            self.out.append(f"// {doc.strip().splitlines()[0]} (`pptx.{module_name(self.src_path)}`)")
        self.out += self.imports()
        for s in self.tree.body:
            self.locals = set()
            if isinstance(s, ast.ClassDef):
                self.out.append("")
                self.out += self.class_def(s)
            elif isinstance(s, ast.FunctionDef):
                self.out.append("")
                self.out += self.jsdoc(s, 0)
                self.locals = set(a.arg for a in s.args.args)
                self.assign_counts = assign_counts(s)
                self.out.append(f"export function {ident(s.name)}({self.params(s, False)}) {{")
                self.out += self.block(s.body, 1)
                self.out.append("}")
            elif isinstance(s, (ast.Assign, ast.AnnAssign)):
                t = s.targets[0] if isinstance(s, ast.Assign) else s.target
                if isinstance(t, ast.Name) and s.value is not None:
                    self.out.append(f"export const {t.id} = {self.expr(s.value)};")
            elif isinstance(s, ast.Expr) and isinstance(s.value, ast.Call):
                self.out.append(self.expr(s.value) + ";")
        text = "\n".join(self.out) + "\n"
        here = module_name(self.src_path).split(".")[:-1]

        def rel(mod):
            p = os.path.relpath(mod, "/".join(here) or ".").replace("\\", "/")
            return p if p.startswith(".") else "./" + p

        helpers = {"defineChildren(": ("defineChildren", "oxml/xmlchemy"), "lazy(this": ("lazy", "util"),
                   "pyInt(": ("pyInt", "util"), "pyFloat(": ("pyFloat", "util"), "pyBool(": ("pyBool", "util"),
                   "pyRound(": ("pyRound", "util"), "pyCmp(": ("pyCmp", "util"), "pyMod(": ("pyMod", "util"), "pyPercent(": ("pyPercent", "util"), "pyRepr(": ("pyRepr", "util"),
                   "parseXml(": ("parseXml", "oxml/parser"), "parseFromTemplate(": ("parseFromTemplate", "oxml/parser")}
        extra = {}
        for exc in ("ValueError", "KeyError", "IndexError", "AttributeError", "NotImplementedError", "ZeroDivisionError", "FileNotFoundError"):
            if re.search(r"\b" + exc + r"\b", text) and not re.search(r"import \{[^}]*\b" + exc + r"\b", text):
                extra.setdefault("errors", []).append(exc)
        for needle, (name, mod) in helpers.items():
            if needle in text and not re.search(r"import \{[^}]*\b" + name + r"\b", text):
                extra.setdefault(mod, []).append(name)
        lines = [f'import {{ {", ".join(v)} }} from "{rel(k)}.js";' for k, v in extra.items()]
        if lines:
            first = text.index("\n") + 1
            text = text[:first] + "\n".join(lines) + "\n" + text[first:]
        return text

    def imports(self):
        here = module_name(self.src_path).split(".")[:-1]
        out = []
        for s in self.tree.body:
            nodes = []
            type_only = False
            if isinstance(s, ast.ImportFrom):
                nodes = [s]
            elif isinstance(s, ast.If) and "TYPE_CHECKING" in ast.unparse(s.test):
                nodes = [x for x in s.body if isinstance(x, ast.ImportFrom)]
                type_only = True
            for imp in nodes:
                if not imp.module or not imp.module.startswith("pptx"):
                    continue
                target = imp.module.split(".")[1:]
                cn = lambda n: n if n[0].isupper() or n.isupper() else camel(n)
                imp_names = [a for a in imp.names if a.name not in ("lazyproperty", "TypedDict")]
                if not imp_names:
                    continue
                names = ", ".join(cn(a.name) if not a.asname else f"{cn(a.name)} as {cn(a.asname)}" for a in imp_names)
                if target == ["oxml"]:
                    target = ["oxml", "parser"]
                elif target and target[-1] in ("shapes", "text", "dml", "chart", "parts", "enum", "opc") and imp.module.count(".") <= 2 and imp.module.split(".")[-1] in ("shapes",):
                    target = target + ["index"]
                path = os.path.relpath("/".join(target) or "index", "/".join(here) or ".").replace("\\", "/")
                if not path.startswith("."):
                    path = "./" + path
                out.append(f'import {"type " if type_only else ""}{{ {names} }} from "{path}.js";')
        return out


def module_name(path):
    p = path.replace("\\", "/")
    p = p[p.index("/pptx/") + 6:]
    return p[:-3].replace("/", ".")


def assign_counts(fn):
    counts = {}
    for n in ast.walk(fn):
        targets = []
        if isinstance(n, ast.Assign):
            targets = n.targets
        elif isinstance(n, (ast.AugAssign, ast.AnnAssign)):
            targets = [n.target]
            if isinstance(n, ast.AugAssign):
                targets = [n.target, n.target]
        elif isinstance(n, ast.For):
            targets = [n.target, n.target]
        for t in targets:
            names = t.elts if isinstance(t, (ast.Tuple, ast.List)) else [t]
            for x in names:
                if isinstance(x, ast.Name):
                    counts[x.id] = counts.get(x.id, 0) + 1
    return counts


def has_yield(fn):
    """True if `fn` itself (not a nested function or lambda) yields."""
    stack = list(ast.iter_child_nodes(fn))
    while stack:
        n = stack.pop()
        if isinstance(n, (ast.Yield, ast.YieldFrom)):
            return True
        if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef, ast.Lambda)):
            continue
        stack.extend(ast.iter_child_nodes(n))
    return False


if __name__ == "__main__":
    src, dst = sys.argv[1], sys.argv[2]
    os.makedirs(os.path.dirname(dst) or ".", exist_ok=True)
    with open(dst, "w", encoding="utf-8", newline="\n") as f:
        f.write(Translator(src).run())
