"""Small LaTeX-math -> Office Math (OMML) converter.

Covers the subset normally found in technical reports: fractions, roots,
sub/superscripts, big operators with limits, \\left..\\right delimiters,
accents, functions, Greek letters, common symbols, \\text, cases/matrices.
Unknown commands are emitted verbatim as upright text instead of failing.
"""

from __future__ import annotations

import re
from xml.sax.saxutils import escape

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

GREEK = {
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ", "epsilon": "ϵ", "varepsilon": "ε",
    "zeta": "ζ", "eta": "η", "theta": "θ", "vartheta": "ϑ", "iota": "ι", "kappa": "κ",
    "lambda": "λ", "mu": "μ", "nu": "ν", "xi": "ξ", "pi": "π", "varpi": "ϖ", "rho": "ρ",
    "varrho": "ϱ", "sigma": "σ", "varsigma": "ς", "tau": "τ", "upsilon": "υ", "phi": "ϕ",
    "varphi": "φ", "chi": "χ", "psi": "ψ", "omega": "ω",
    "Gamma": "Γ", "Delta": "Δ", "Theta": "Θ", "Lambda": "Λ", "Xi": "Ξ", "Pi": "Π",
    "Sigma": "Σ", "Upsilon": "Υ", "Phi": "Φ", "Psi": "Ψ", "Omega": "Ω",
}

SYMBOLS = {
    "cdot": "·", "times": "×", "div": "÷", "pm": "±", "mp": "∓", "ast": "∗", "star": "⋆",
    "circ": "∘", "bullet": "•", "le": "≤", "leq": "≤", "ge": "≥", "geq": "≥", "ne": "≠",
    "neq": "≠", "approx": "≈", "equiv": "≡", "sim": "∼", "simeq": "≃", "cong": "≅",
    "propto": "∝", "ll": "≪", "gg": "≫", "infty": "∞", "to": "→", "rightarrow": "→",
    "leftarrow": "←", "Rightarrow": "⇒", "Leftarrow": "⇐", "Leftrightarrow": "⇔",
    "leftrightarrow": "↔", "mapsto": "↦", "implies": "⇒", "iff": "⇔", "partial": "∂",
    "nabla": "∇", "ldots": "…", "dots": "…", "cdots": "⋯", "vdots": "⋮", "ddots": "⋱",
    "in": "∈", "notin": "∉", "ni": "∋", "subset": "⊂", "subseteq": "⊆", "supset": "⊃",
    "supseteq": "⊇", "cup": "∪", "cap": "∩", "setminus": "∖", "emptyset": "∅",
    "varnothing": "∅", "forall": "∀", "exists": "∃", "neg": "¬", "land": "∧", "lor": "∨",
    "wedge": "∧", "vee": "∨", "oplus": "⊕", "otimes": "⊗", "perp": "⊥", "parallel": "∥",
    "angle": "∠", "triangle": "△", "degree": "°", "prime": "′", "hbar": "ℏ", "ell": "ℓ",
    "Re": "ℜ", "Im": "ℑ", "aleph": "ℵ", "lbrace": "{", "rbrace": "}", "langle": "⟨",
    "rangle": "⟩", "lfloor": "⌊", "rfloor": "⌋", "lceil": "⌈", "rceil": "⌉", "vert": "|",
    "Vert": "‖", "mid": "|", "colon": ":", "cdotp": "·", "dagger": "†",
}

NARY = {
    "sum": "∑", "prod": "∏", "coprod": "∐", "int": "∫", "iint": "∬", "iiint": "∭",
    "oint": "∮", "bigcup": "⋃", "bigcap": "⋂", "bigoplus": "⨁", "bigotimes": "⨂",
}

FUNCS = {
    "sin", "cos", "tan", "cot", "sec", "csc", "tg", "ctg", "arcsin", "arccos", "arctan",
    "arctg", "arcctg", "sinh", "cosh", "tanh", "coth", "sh", "ch", "th", "cth", "ln", "lg",
    "log", "exp", "det", "dim", "ker", "deg", "gcd", "arg", "max", "min", "lim", "sup",
    "inf", "limsup", "liminf", "Pr",
}
LIMIT_FUNCS = {"max", "min", "lim", "sup", "inf", "limsup", "liminf"}

ACCENTS = {
    "hat": "\u0302", "widehat": "\u0302", "bar": "\u0305", "vec": "\u20d7",
    "overrightarrow": "\u20d7", "tilde": "\u0303", "widetilde": "\u0303", "dot": "\u0307",
    "ddot": "\u0308", "check": "\u030c", "breve": "\u0306", "acute": "\u0301", "grave": "\u0300",
}

SPACES = {",": "\u2009", ":": "\u205f", ">": "\u205f", ";": "\u2004", "quad": "\u2003",
          "qquad": "\u2003\u2003", " ": " ", "enspace": "\u2002", "!": ""}

TEXT_CMDS = {"text", "textrm", "mbox", "mathrm", "operatorname", "textnormal", "textit",
             "textbf", "mathbf", "mathit", "boldsymbol", "mathsf", "mathtt", "texttt"}

DELIMS = {"(": "(", ")": ")", "[": "[", "]": "]", "\\{": "{", "\\}": "}", "|": "|",
          "\\|": "‖", "\\lbrace": "{", "\\rbrace": "}", "\\langle": "⟨", "\\rangle": "⟩",
          "\\lfloor": "⌊", "\\rfloor": "⌋", "\\lceil": "⌈", "\\rceil": "⌉", "\\vert": "|",
          "\\Vert": "‖", "\\lvert": "|", "\\rvert": "|", ".": "", "/": "/"}

RELATIONS = set("=<>≤≥≠≈≡∼→←⇒⇔∈+−±∓,;")

ENV_DELIMS = {"matrix": ("", ""), "pmatrix": ("(", ")"), "bmatrix": ("[", "]"),
              "Bmatrix": ("{", "}"), "vmatrix": ("|", "|"), "Vmatrix": ("‖", "‖"),
              "cases": ("{", ""), "array": ("", ""), "aligned": None, "align": None,
              "align*": None, "gathered": None, "gather": None, "gather*": None, "split": None,
              "eqnarray": None}


class LatexError(ValueError):
    pass


# --------------------------------------------------------------------------- tokenizer

_TOKEN = re.compile(r"\\([A-Za-z]+\*?)|\\(.)|(\s+)|(.)", re.S)


def _tokenize(src: str):
    toks = []
    for m in _TOKEN.finditer(src):
        if m.group(1) is not None:
            toks.append(("cmd", m.group(1)))
        elif m.group(2) is not None:
            toks.append(("cmd", m.group(2)))
        elif m.group(3) is not None:
            toks.append(("ws", m.group(3)))
        else:
            toks.append(("chr", m.group(4)))
    return toks


# --------------------------------------------------------------------------- parser
# Nodes: ("run", text, kind) kind in var|num|op|text|fn|bold
#        ("group", [nodes]) ("frac", n, d) ("sqrt", deg, body) ("scr", base, sub, sup)
#        ("delim", l, r, body) ("nary", chr, sub, sup, body) ("acc", chr, body)
#        ("bar", pos, body) ("func", name_node, body) ("limlow", base, lim)
#        ("matrix", rows, l, r) ("eqarr", rows)


class _Parser:
    def __init__(self, src: str):
        self.toks = _tokenize(src)
        self.i = 0

    def peek(self, skip_ws=True):
        j = self.i
        while skip_ws and j < len(self.toks) and self.toks[j][0] == "ws":
            j += 1
        return self.toks[j] if j < len(self.toks) else None

    def next(self, skip_ws=True):
        while skip_ws and self.i < len(self.toks) and self.toks[self.i][0] == "ws":
            self.i += 1
        if self.i >= len(self.toks):
            return None
        t = self.toks[self.i]
        self.i += 1
        return t

    # stop: set of tokens ("chr","}") / ("cmd","right") / ("chr","&") / ("cmd","\\")
    def parse_list(self, stop=frozenset()):
        nodes = []
        while True:
            t = self.peek()
            if t is None:
                if ("chr", "}") in stop:
                    raise LatexError("missing '}'")
                return nodes, None
            if t in stop:
                self.next()
                return nodes, t
            if t == ("chr", "^") or t == ("chr", "_"):
                self.next()
                self._attach_script(nodes, t[1])
                continue
            if t == ("cmd", "limits") or t == ("cmd", "nolimits"):
                self.next()
                continue
            node = self.parse_atom(top=True)
            if node is not None:
                nodes.append(node)
        # unreachable

    def _attach_script(self, nodes, kind):
        arg = self.parse_atom()
        if arg is None:
            raise LatexError(f"missing argument for '{kind}'")
        base = nodes.pop() if nodes else ("run", "", "text")
        if base[0] == "nary" and base[4] is None:
            _, ch, sub, sup, body = base
            if kind == "_":
                sub = arg
            else:
                sup = arg
            nodes.append(("nary", ch, sub, sup, None))
            return
        if base[0] == "run" and base[2] == "fn" and base[1] in LIMIT_FUNCS and kind == "_":
            nodes.append(("limlow", base, arg))
            return
        if base[0] == "scr":
            _, b, sub, sup = base
            if kind == "_" and sub is None:
                nodes.append(("scr", b, arg, sup))
                return
            if kind == "^" and sup is None:
                nodes.append(("scr", b, sub, arg))
                return
        if kind == "_":
            nodes.append(("scr", base, arg, None))
        else:
            nodes.append(("scr", base, None, arg))

    def parse_arg(self):
        node = self.parse_atom()
        if node is None:
            raise LatexError("missing argument")
        return node

    def raw_group(self) -> str:
        """Read a braced argument verbatim (for \\text)."""
        t = self.next()
        if t != ("chr", "{"):
            return _tok_text(t)
        depth, out = 1, []
        while True:
            t = self.next(skip_ws=False)
            if t is None:
                raise LatexError("missing '}'")
            if t == ("chr", "{"):
                depth += 1
            elif t == ("chr", "}"):
                depth -= 1
                if depth == 0:
                    return "".join(out)
            out.append(_tok_text(t))

    def parse_atom(self, top=False):
        t = self.next()
        if t is None:
            return None
        kind, val = t
        if kind == "chr":
            if val == "{":
                nodes, _ = self.parse_list(stop=frozenset({("chr", "}")}))
                return ("group", nodes)
            if val == "}":
                raise LatexError("unexpected '}'")
            if val == "&":
                return ("run", " ", "op")
            if val == "~":
                return ("run", "\u00a0", "op")
            if val.isdigit() and top:
                return self._number(val)
            return _char_node(val)
        # commands
        if val in ("frac", "dfrac", "tfrac", "cfrac"):
            return ("frac", self.parse_arg(), self.parse_arg())
        if val == "binom":
            return ("delim", "(", ")", [("frac", self.parse_arg(), self.parse_arg(), "nobar")])
        if val == "sqrt":
            deg = None
            if self.peek() == ("chr", "["):
                self.next()
                nodes, _ = self.parse_list(stop=frozenset({("chr", "]")}))
                deg = ("group", nodes)
            return ("sqrt", deg, self.parse_arg())
        if val == "left":
            left = self._delim()
            nodes, _ = self.parse_list(stop=frozenset({("cmd", "right")}))
            right = self._delim()
            return ("delim", left, right, nodes)
        if val in ("big", "Big", "bigg", "Bigg", "bigl", "bigr", "Bigl", "Bigr"):
            return ("run", self._delim(), "op")
        if val in TEXT_CMDS:
            if val in ("mathbf", "boldsymbol", "textbf"):
                inner = self.parse_arg()
                return ("bold", inner)
            if val == "operatorname":
                return ("run", self.raw_group(), "fn")
            return ("run", self.raw_group(), "text")
        if val in ("mathit",):
            return self.parse_arg()
        if val in GREEK:
            return ("run", GREEK[val], "var")
        if val in NARY:
            return ("nary", NARY[val], None, None, None)
        if val in FUNCS:
            return ("run", val, "fn")
        if val in ACCENTS:
            return ("acc", ACCENTS[val], self.parse_arg())
        if val in ("overline",):
            return ("bar", "top", self.parse_arg())
        if val in ("underline",):
            return ("bar", "bot", self.parse_arg())
        if val in SPACES:
            return ("run", SPACES[val], "op")
        if val in SYMBOLS:
            return ("run", SYMBOLS[val], "op")
        if val in ("{", "}", "%", "$", "_", "&", "#", "|"):
            return ("run", "‖" if val == "|" else val, "op")
        if val == "\\":
            return ("run", " ", "op")
        if val == "begin":
            return self._env()
        if val in ("displaystyle", "textstyle", "scriptstyle", "nonumber", "notag"):
            return None
        if val in ("mathbb", "mathcal", "mathfrak", "mathscr"):
            return ("run", self.raw_group(), "var")
        return ("run", "\\" + val, "text")

    def _number(self, first):
        out = first
        while True:
            t = self.peek(skip_ws=False)
            if t and t[0] == "chr" and t[1].isdigit():
                out += t[1]
                self.next(skip_ws=False)
            elif t and t[0] == "chr" and t[1] in ".," and self._digit_after():
                out += t[1]
                self.next(skip_ws=False)
            elif t == ("chr", "{") and self._is_brace_comma():
                self.i += 3
                out += ","
            else:
                return ("run", out, "num")

    def _digit_after(self):
        j = self.i + 1
        return j < len(self.toks) and self.toks[j][0] == "chr" and self.toks[j][1].isdigit()

    def _is_brace_comma(self):
        t = self.toks[self.i:self.i + 4]
        return (len(t) == 4 and t[1] == ("chr", ",") and t[2] == ("chr", "}")
                and t[3][0] == "chr" and t[3][1].isdigit())

    def _delim(self):
        t = self.next()
        if t is None:
            raise LatexError("missing delimiter")
        key = t[1] if t[0] == "chr" else "\\" + t[1]
        if key not in DELIMS:
            raise LatexError(f"unknown delimiter {key!r}")
        return DELIMS[key]

    def _env(self):
        name = self.raw_group()
        if name not in ENV_DELIMS:
            raise LatexError(f"unsupported environment {name!r}")
        if name == "array" and self.peek() == ("chr", "{"):
            self.raw_group()  # column spec
        rows, row = [], []
        stop = frozenset({("chr", "&"), ("cmd", "\\"), ("cmd", "end")})
        while True:
            nodes, t = self.parse_list(stop=stop)
            row.append(nodes)
            if t is None:
                raise LatexError(f"missing \\end{{{name}}}")
            if t == ("cmd", "\\"):
                rows.append(row)
                row = []
            elif t == ("cmd", "end"):
                self.raw_group()
                if any(c for c in row):
                    rows.append(row)
                break
        delims = ENV_DELIMS[name]
        if delims is None:
            return ("eqarr", rows)
        if name == "cases":
            return ("delim", "{", "", [("eqarr", rows)])
        return ("matrix", rows, delims[0], delims[1])


def _tok_text(t):
    if t is None:
        return ""
    if t[0] == "cmd":
        return SPACES.get(t[1], SYMBOLS.get(t[1], GREEK.get(t[1], "\\" + t[1] if len(t[1]) > 1 else t[1])))
    return t[1]


def _char_node(ch):
    if ch == "-":
        return ("run", "−", "op")
    if ch == "*":
        return ("run", "∗", "op")
    if ch == "'":
        return ("run", "′", "op")
    if ch.isdigit():
        return ("run", ch, "num")
    if "a" <= ch.lower() <= "z":
        return ("run", ch, "var")
    if ch.isalpha():  # Cyrillic and other scripts: upright, as in Russian technical texts
        return ("run", ch, "text")
    return ("run", ch, "op")


def _attach_bodies(nodes):
    """Give big operators and functions their operand (what follows them)."""
    out = []
    i = 0
    nodes = [_walk(n) for n in nodes]
    while i < len(nodes):
        n = nodes[i]
        if n[0] == "nary" and n[4] is None:
            j = i + 1
            body = []
            while j < len(nodes) and not _is_relation(nodes[j]):
                body.append(nodes[j])
                j += 1
            out.append(("nary", n[1], n[2], n[3], body))
            i = j
            continue
        if (n[0] == "run" and n[2] == "fn") or n[0] == "limlow":
            if i + 1 < len(nodes) and not _is_relation(nodes[i + 1]):
                out.append(("func", n, [nodes[i + 1]]))
                i += 2
                continue
        out.append(n)
        i += 1
    return out


def _is_relation(n):
    return n[0] == "run" and n[2] == "op" and n[1] and n[1][0] in RELATIONS


def _walk(n):
    t = n[0]
    if t == "group":
        return ("group", _attach_bodies(n[1]))
    if t == "delim":
        return ("delim", n[1], n[2], _attach_bodies(n[3]))
    if t == "frac":
        return ("frac", _walk(n[1]), _walk(n[2])) + tuple(n[3:])
    if t == "sqrt":
        return ("sqrt", _walk(n[1]) if n[1] else None, _walk(n[2]))
    if t == "scr":
        return ("scr", _walk(n[1]), _walk(n[2]) if n[2] else None, _walk(n[3]) if n[3] else None)
    if t in ("acc", "bar"):
        return (t, n[1], _walk(n[2]))
    if t == "bold":
        return ("bold", _walk(n[1]))
    if t == "limlow":
        return ("limlow", n[1], _walk(n[2]))
    if t == "matrix":
        return ("matrix", [[_attach_bodies(c) for c in r] for r in n[1]], n[2], n[3])
    if t == "eqarr":
        return ("eqarr", [[_attach_bodies(c) for c in r] for r in n[1]])
    if t == "nary":
        return ("nary", n[1], _walk(n[2]) if n[2] else None, _walk(n[3]) if n[3] else None, n[4])
    return n


def parse(src: str):
    p = _Parser(src)
    stop = frozenset({("cmd", "\\")})
    rows = []
    while True:
        nodes, t = p.parse_list(stop=stop)
        rows.append(nodes)
        if t is None:
            break
    rows = [_attach_bodies(r) for r in rows if r or len(rows) == 1]
    if len(rows) > 1:  # top-level line breaks -> equation array
        return [("eqarr", [[r] for r in rows])]
    return rows[0] if rows else []


# --------------------------------------------------------------------------- emitter


class _Emitter:
    def __init__(self, italic: bool, text_font: str):
        self.italic = italic
        self.text_font = text_font

    def run(self, text, kind, bold=False):
        if text == "":
            return ""
        rpr = []
        if kind == "text":
            rpr.append("<m:nor/>")
        elif kind == "var" and self.italic:
            if bold:
                rpr.append('<m:sty m:val="bi"/>')
        else:
            rpr.append('<m:sty m:val="%s"/>' % ("b" if bold else "p"))
        font = self.text_font if kind == "text" else "Cambria Math"
        wrpr = '<w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s" w:cs="%s"/>%s</w:rPr>' % (
            font, font, font, "<w:b/>" if bold and kind == "text" else "")
        return '<m:r><m:rPr>%s</m:rPr>%s<m:t xml:space="preserve">%s</m:t></m:r>' % (
            "".join(rpr), wrpr, escape(text))

    def seq(self, nodes, bold=False):
        out = []
        buf_text, buf_kind = "", None
        for n in nodes:
            if n[0] == "run" and n[2] != "fn":
                kind = n[2]
                if kind == buf_kind:
                    buf_text += n[1]
                    continue
                out.append(self.run(buf_text, buf_kind, bold) if buf_kind else "")
                buf_text, buf_kind = n[1], kind
                continue
            if buf_kind:
                out.append(self.run(buf_text, buf_kind, bold))
                buf_text, buf_kind = "", None
            out.append(self.node(n, bold))
        if buf_kind:
            out.append(self.run(buf_text, buf_kind, bold))
        return "".join(out)

    def e(self, n, bold=False, tag="m:e"):
        if n is None:
            return "<%s/>" % tag
        nodes = n[1] if n[0] == "group" else (n if isinstance(n, list) else [n])
        return "<%s>%s</%s>" % (tag, self.seq(nodes, bold), tag)

    def node(self, n, bold=False):
        t = n[0]
        if t == "run":
            return self.run(n[1], "fn" if n[2] == "fn" else n[2], bold)
        if t == "group":
            return self.seq(n[1], bold)
        if t == "bold":
            return self.e_inline(n[1], True)
        if t == "frac":
            pr = '<m:fPr><m:type m:val="noBar"/></m:fPr>' if len(n) > 3 else ""
            return "<m:f>%s%s%s</m:f>" % (pr, self.e(n[1], bold, "m:num"), self.e(n[2], bold, "m:den"))
        if t == "sqrt":
            if n[1] is None:
                return '<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/>%s</m:rad>' % self.e(n[2], bold)
            return "<m:rad>%s%s</m:rad>" % (self.e(n[1], bold, "m:deg"), self.e(n[2], bold))
        if t == "scr":
            base, sub, sup = n[1], n[2], n[3]
            if sub is not None and sup is not None:
                return "<m:sSubSup>%s%s%s</m:sSubSup>" % (
                    self.e(base, bold), self.e(sub, bold, "m:sub"), self.e(sup, bold, "m:sup"))
            if sub is not None:
                return "<m:sSub>%s%s</m:sSub>" % (self.e(base, bold), self.e(sub, bold, "m:sub"))
            return "<m:sSup>%s%s</m:sSup>" % (self.e(base, bold), self.e(sup, bold, "m:sup"))
        if t == "delim":
            return ('<m:d><m:dPr><m:begChr m:val="%s"/><m:endChr m:val="%s"/></m:dPr>%s</m:d>'
                    % (escape(n[1], {'"': "&quot;"}), escape(n[2], {'"': "&quot;"}),
                       self.e(n[3], bold)))
        if t == "nary":
            ch, sub, sup, body = n[1], n[2], n[3], n[4] or []
            loc = "subSup" if ch in "∫∬∭∮" else "undOvr"
            pr = '<m:naryPr><m:chr m:val="%s"/><m:limLoc m:val="%s"/>%s%s</m:naryPr>' % (
                ch, loc, '<m:subHide m:val="1"/>' if sub is None else "",
                '<m:supHide m:val="1"/>' if sup is None else "")
            return "<m:nary>%s%s%s%s</m:nary>" % (
                pr, self.e(sub, bold, "m:sub"), self.e(sup, bold, "m:sup"), self.e(body, bold))
        if t == "acc":
            return '<m:acc><m:accPr><m:chr m:val="%s"/></m:accPr>%s</m:acc>' % (n[1], self.e(n[2], bold))
        if t == "bar":
            return '<m:bar><m:barPr><m:pos m:val="%s"/></m:barPr>%s</m:bar>' % (n[1], self.e(n[2], bold))
        if t == "limlow":
            return "<m:limLow>%s%s</m:limLow>" % (self.e(n[1], bold), self.e(n[2], bold, "m:lim"))
        if t == "func":
            return "<m:func><m:fName>%s</m:fName>%s</m:func>" % (self.node(n[1], bold), self.e(n[2], bold))
        if t == "matrix":
            cols = max(len(r) for r in n[1])
            mpr = ('<m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val="%d"/><m:mcJc m:val="center"/>'
                   '</m:mcPr></m:mc></m:mcs></m:mPr>' % cols)
            rows = "".join("<m:mr>%s</m:mr>" % "".join(self.e(c, bold) for c in r + [[]] * (cols - len(r)))
                           for r in n[1])
            m = "<m:m>%s%s</m:m>" % (mpr, rows)
            if n[2] or n[3]:
                return ('<m:d><m:dPr><m:begChr m:val="%s"/><m:endChr m:val="%s"/></m:dPr><m:e>%s</m:e></m:d>'
                        % (n[2], n[3], m))
            return m
        if t == "eqarr":
            rows = []
            for r in n[1]:
                # '&' alignment points become OMML alignment markers
                rows.append("<m:e>%s</m:e>" % self._aligned([self.seq(c, bold) for c in r]))
            return '<m:eqArr><m:eqArrPr><m:baseJc m:val="center"/></m:eqArrPr>%s</m:eqArr>' % "".join(rows)
        raise LatexError(f"internal: unknown node {t}")

    def _aligned(self, cells):
        if len(cells) == 1:
            return cells[0]
        aln = '<m:r><m:rPr><m:aln/></m:rPr><m:t></m:t></m:r>'
        return aln.join(cells)

    def e_inline(self, n, bold):
        nodes = n[1] if n[0] == "group" else [n]
        return self.seq(nodes, bold)


def to_omml(latex: str, italic: bool = False, text_font: str = "Times New Roman") -> str:
    """Return the inner content of an <m:oMath> element for the given LaTeX source."""
    nodes = parse(latex.strip())
    return _Emitter(italic, text_font).seq(nodes)


def omath(latex: str, italic: bool = False, text_font: str = "Times New Roman") -> str:
    return '<m:oMath xmlns:m="%s" xmlns:w="%s">%s</m:oMath>' % (
        M_NS, W_NS, to_omml(latex, italic, text_font))


def omath_para(latex: str, italic: bool = False, text_font: str = "Times New Roman") -> str:
    return ('<m:oMathPara xmlns:m="%s" xmlns:w="%s"><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr>'
            '<m:oMath>%s</m:oMath></m:oMathPara>' % (M_NS, W_NS, to_omml(latex, italic, text_font)))
