"""Markdown -> list of document blocks (intermediate representation).

The IR is deliberately small; numbering and rendering happen later in
``numbering.py`` and ``render.py``.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

import yaml
from markdown_it import MarkdownIt
from mdit_py_plugins.dollarmath import dollarmath_plugin
from mdit_py_plugins.front_matter import front_matter_plugin

# ----------------------------------------------------------------------------- inline IR
# An inline is a tuple:
#   ("text", str, fmt)   fmt: frozenset of {"b","i","s","code","sub","sup"}
#   ("math", latex)
#   ("ref", "fig:x")     resolved during rendering
#   ("br",)


@dataclass
class Heading:
    level: int
    inlines: list
    id: str | None = None
    unnumbered: bool = False
    appendix: bool = False
    appendix_letter: str | None = None
    appendix_status: str | None = None
    number: str | None = None      # filled by numbering


@dataclass
class Paragraph:
    inlines: list
    kind: str = "body"             # body | quote


@dataclass
class Figure:
    src: str
    caption: list
    id: str | None = None
    width: str | None = None
    number: str | None = None


@dataclass
class Table:
    rows: list                     # list of rows; row = list of cells; cell = inlines
    aligns: list
    header_rows: int = 1
    caption: list | None = None
    id: str | None = None
    widths: list | None = None
    number: str | None = None


@dataclass
class Code:
    lines: list
    lang: str = ""
    caption: list | None = None
    id: str | None = None
    number: str | None = None


@dataclass
class Math:
    latex: str
    id: str | None = None
    numbered: bool = True
    number: str | None = None


@dataclass
class Where:
    lines: list                    # list of inline lists; first starts with "где"


@dataclass
class ListBlock:
    ordered: bool
    items: list                    # list of lists of blocks
    start: int = 1
    sources: bool = False
    item_ids: list = field(default_factory=list)


@dataclass
class PageBreak:
    pass


class MarkdownError(ValueError):
    pass


# ----------------------------------------------------------------------------- helpers

_ATTR_TAIL = re.compile(r"\s*\{([^{}]*)\}\s*$")
# Values may be quoted with "", '' or, after smart-quote typography ran, «» / „“.
_ATTR_ITEM = re.compile(r'([#.]?[\w:\-.]+)(?:=("([^"]*)"|\'([^\']*)\'|«([^»]*)»|„([^“]*)“|[^\s"\'«„]+))?|(-)')
_CAPTION = re.compile(r"^\s*(Table|Таблица|Listing|Листинг|Code|Код)\s*:\s*", re.I)
_APPENDIX = re.compile(
    r"^\s*(?:ПРИЛОЖЕНИЕ|Приложение|APPENDIX|Appendix)\b\s*([А-ЯA-Z](?![\w]))?\s*(?:\(([^)]*)\))?\s*(.*)$")
_PAGEBREAK = re.compile(r"^\s*(\\newpage|\\pagebreak|<!--\s*(pagebreak|newpage)\s*-->)\s*$", re.I)
_SPACED_HYPHEN = re.compile(r"(?<=\S) - (?=\S)")

STRUCTURAL = {
    "ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "РЕФЕРАТ", "АННОТАЦИЯ", "СОДЕРЖАНИЕ", "ОГЛАВЛЕНИЕ",
    "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", "СПИСОК ИСПОЛЬЗОВАННОЙ ЛИТЕРАТУРЫ", "СПИСОК ЛИТЕРАТУРЫ",
    "БИБЛИОГРАФИЧЕСКИЙ СПИСОК", "ТЕРМИНЫ И ОПРЕДЕЛЕНИЯ", "ОПРЕДЕЛЕНИЯ",
    "ПЕРЕЧЕНЬ СОКРАЩЕНИЙ И ОБОЗНАЧЕНИЙ", "ПЕРЕЧЕНЬ СОКРАЩЕНИЙ", "СПИСОК СОКРАЩЕНИЙ",
    "ОБОЗНАЧЕНИЯ И СОКРАЩЕНИЯ", "СПИСОК ИСПОЛНИТЕЛЕЙ",
}
SOURCES_TITLES = {"СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", "СПИСОК ИСПОЛЬЗОВАННОЙ ЛИТЕРАТУРЫ",
                  "СПИСОК ЛИТЕРАТУРЫ", "БИБЛИОГРАФИЧЕСКИЙ СПИСОК"}

# label names may contain - and . inside, but not at the end ("@fig:a." is a ref + full stop)
REF_RE = re.compile(r"@((?:fig|tbl|lst|eq|sec|app|src):\w(?:[\w\-.]*\w)?)")


def parse_attrs(s: str) -> dict:
    out: dict = {"classes": []}
    for m in _ATTR_ITEM.finditer(s):
        if m.group(7):
            out["unnumbered"] = True
            continue
        key = m.group(1)
        if key.startswith("#"):
            out["id"] = key[1:]
        elif key.startswith("."):
            out["classes"].append(key[1:])
            if key == ".unnumbered":
                out["unnumbered"] = True
        elif m.group(2) is not None:
            quoted = [g for g in m.group(3, 4, 5, 6) if g is not None]
            out[key] = quoted[0] if quoted else m.group(2)
    return out


def split_attrs(text: str):
    m = _ATTR_TAIL.search(text)
    if not m:
        return text, {}
    return text[: m.start()], parse_attrs(m.group(1))


def plain(inlines) -> str:
    out = []
    for it in inlines:
        if it[0] == "text":
            out.append(it[1])
        elif it[0] == "math":
            out.append(it[1])
        elif it[0] == "ref":
            out.append("@" + it[1])
        elif it[0] in ("br", "soft"):
            out.append(" ")
    return "".join(out)


def _strip_tail_attrs(inlines):
    """Remove a trailing ``{...}`` attribute block from inline list."""
    if not inlines or inlines[-1][0] != "text":
        return inlines, {}
    text, attrs = split_attrs(inlines[-1][1])
    if not attrs:
        return inlines, {}
    rest = list(inlines[:-1])
    if text.strip():
        rest.append(("text", text.rstrip(), inlines[-1][2]))
    elif rest and rest[-1][0] == "text":
        rest[-1] = ("text", rest[-1][1].rstrip(), rest[-1][2])
    return rest, attrs


def _strip_prefix(inlines, regex):
    if not inlines or inlines[0][0] != "text":
        return inlines
    t = regex.sub("", inlines[0][1], count=1)
    return ([("text", t, inlines[0][2])] if t else []) + list(inlines[1:])


# ----------------------------------------------------------------------------- parser


def make_md(typography: bool) -> MarkdownIt:
    md = MarkdownIt("commonmark", {"typographer": typography, "html": True})
    md.options["quotes"] = "«»„“"
    md.enable(["table", "strikethrough"])
    if typography:
        md.enable(["replacements", "smartquotes"])
    md.use(front_matter_plugin)
    md.use(dollarmath_plugin, allow_space=True, allow_digits=True, double_inline=True)
    return md


_EQ_ATTR = re.compile(r"\$\$[ \t]*\{#(eq:[^}\s]+)\}")
_EQ_NONUM = re.compile(r"\$\$[ \t]*\{-\}")


def parse_document(src: str, typography: bool = True):
    """Return (front_matter_dict, blocks)."""
    src = src.replace("\r\n", "\n").lstrip("\ufeff")
    src = _EQ_ATTR.sub(r"$$ (\1)", src)
    src = _EQ_NONUM.sub("$$ (-)", src)
    md = make_md(typography)
    tokens = md.parse(src)
    meta: dict = {}
    if tokens and tokens[0].type == "front_matter":
        meta = yaml.safe_load(tokens[0].content) or {}
        if not isinstance(meta, dict):
            raise MarkdownError("front matter must be a YAML mapping")
        tokens = tokens[1:]
    conv = _Converter(typography)
    blocks = conv.blocks(tokens, 0, len(tokens))
    blocks = _postprocess(blocks)
    return meta, blocks


class _Converter:
    def __init__(self, typography: bool):
        self.typography = typography

    # -- block level ------------------------------------------------------------------
    def blocks(self, toks, i, end):
        out = []
        while i < end:
            t = toks[i]
            tp = t.type
            if tp == "heading_open":
                inl = self.inlines(toks[i + 1].children)
                out.append(self.heading(int(t.tag[1]), inl))
                i += 3
            elif tp == "paragraph_open":
                close = _find_close(toks, i)
                inl_tok = toks[i + 1]
                out.extend(self.paragraph(inl_tok))
                i = close + 1
            elif tp in ("bullet_list_open", "ordered_list_open"):
                close = _find_close(toks, i)
                out.append(self.list_block(toks, i, close))
                i = close + 1
            elif tp == "table_open":
                close = _find_close(toks, i)
                out.append(self.table(toks, i, close))
                i = close + 1
            elif tp in ("fence", "code_block"):
                out.append(self.code(t))
                i += 1
            elif tp in ("math_block", "math_block_label"):
                label = t.info.strip() if tp == "math_block_label" else ""
                numbered = label != "-"
                out.append(Math(t.content.strip(), id=label if label and label != "-" else None,
                                numbered=numbered))
                i += 1
            elif tp == "blockquote_open":
                close = _find_close(toks, i)
                for b in self.blocks(toks, i + 1, close):
                    if isinstance(b, Paragraph):
                        b.kind = "quote"
                    out.append(b)
                i = close + 1
            elif tp == "hr":
                out.append(PageBreak())
                i += 1
            elif tp == "html_block":
                if _PAGEBREAK.match(t.content.strip()):
                    out.append(PageBreak())
                i += 1
            else:
                i += 1
        return out

    def heading(self, level, inl):
        inl, attrs = _strip_tail_attrs(inl)
        h = Heading(level, inl, id=attrs.get("id"), unnumbered=bool(attrs.get("unnumbered")))
        text = plain(inl).strip()
        if level == 1:
            m = _APPENDIX.match(text)
            if m:
                h.appendix = True
                h.appendix_letter = m.group(1)
                h.appendix_status = (m.group(2) or "").strip() or None
                h.inlines = [("text", m.group(3).strip(), frozenset())] if m.group(3).strip() else []
            elif text.upper() in STRUCTURAL:
                h.unnumbered = True
        return h

    def paragraph(self, inl_tok):
        raw = inl_tok.content.strip()
        if _PAGEBREAK.match(raw):
            return [PageBreak()]
        children = inl_tok.children or []
        # figure: paragraph made of a single image (+ optional {attrs})
        imgs = [c for c in children if c.type == "image"]
        others = [c for c in children if c.type != "image" and not (c.type == "text" and not c.content.strip())
                  and c.type != "softbreak"]
        if len(imgs) == 1 and all(c.type == "text" and _ATTR_TAIL.fullmatch(c.content) for c in others):
            img = imgs[0]
            attrs = parse_attrs(_ATTR_TAIL.fullmatch(others[0].content).group(1)) if others else {}
            caption = self.inlines(img.children or [])
            return [Figure(img.attrGet("src"), caption, id=attrs.get("id"), width=attrs.get("width"))]
        inl = self.inlines(children)
        return [Paragraph(inl)]

    def list_block(self, toks, i, close):
        ordered = toks[i].type == "ordered_list_open"
        start = int(toks[i].attrGet("start") or 1)
        items, ids = [], []
        j = i + 1
        while j < close:
            if toks[j].type == "list_item_open":
                c = _find_close(toks, j)
                blocks = self.blocks(toks, j + 1, c)
                item_id = None
                if blocks and isinstance(blocks[-1], Paragraph):
                    inl, attrs = _strip_tail_attrs(blocks[-1].inlines)
                    if attrs.get("id"):
                        blocks[-1].inlines = inl
                        item_id = attrs["id"]
                items.append(blocks)
                ids.append(item_id)
                j = c + 1
            else:
                j += 1
        return ListBlock(ordered, items, start=start, item_ids=ids)

    def table(self, toks, i, close):
        rows, aligns, header_rows = [], [], 0
        row = None
        in_head = False
        for j in range(i, close):
            t = toks[j]
            if t.type == "thead_open":
                in_head = True
            elif t.type == "thead_close":
                in_head = False
            elif t.type == "tr_open":
                row = []
            elif t.type == "tr_close":
                rows.append(row)
                if in_head:
                    header_rows += 1
            elif t.type in ("th_open", "td_open"):
                style = t.attrGet("style") or ""
                m = re.search(r"text-align:(\w+)", style)
                if len(rows) == 0:
                    aligns.append(m.group(1) if m else None)
            elif t.type == "inline":
                row.append(self.inlines(t.children or []))
        return Table(rows, aligns, header_rows=header_rows)

    def code(self, t):
        info = t.info.strip() if t.type == "fence" else ""
        info, attrs = split_attrs(info)
        lang = info.split()[0] if info.split() else ""
        lines = t.content.rstrip("\n").split("\n")
        caption = None
        if attrs.get("caption"):
            caption = [("text", attrs["caption"], frozenset())]
        return Code(lines, lang=lang, caption=caption, id=attrs.get("id"))

    # -- inline level -----------------------------------------------------------------
    def inlines(self, children):
        out = []
        fmt: set = set()
        link_stack = []
        for c in children:
            tp = c.type
            if tp == "text":
                self._text(out, c.content, fmt)
            elif tp == "softbreak":
                out.append(("soft",))
            elif tp == "hardbreak":
                out.append(("br",))
            elif tp == "code_inline":
                out.append(("text", c.content, frozenset(fmt | {"code"})))
            elif tp in ("math_inline", "math_inline_double"):
                out.append(("math", c.content.strip()))
            elif tp == "em_open":
                fmt.add("i")
            elif tp == "em_close":
                fmt.discard("i")
            elif tp == "strong_open":
                fmt.add("b")
            elif tp == "strong_close":
                fmt.discard("b")
            elif tp == "s_open":
                fmt.add("s")
            elif tp == "s_close":
                fmt.discard("s")
            elif tp == "link_open":
                link_stack.append((c.attrGet("href"), len(out)))
            elif tp == "link_close":
                href, pos = link_stack.pop()
                label = plain(out[pos:]).strip()
                if not label:
                    out.append(("text", href, frozenset(fmt)))
            elif tp == "image":
                self._text(out, "".join(x.content for x in (c.children or [])), fmt)
            elif tp == "html_inline":
                tag = c.content.strip().lower()
                if tag in ("<br>", "<br/>", "<br />"):
                    out.append(("br",))
                elif tag in ("<sub>", "<sup>", "<b>", "<i>", "<u>"):
                    fmt.add({"<sub>": "sub", "<sup>": "sup", "<b>": "b", "<i>": "i", "<u>": "u"}[tag])
                elif tag in ("</sub>", "</sup>", "</b>", "</i>", "</u>"):
                    fmt.discard({"</sub>": "sub", "</sup>": "sup", "</b>": "b", "</i>": "i", "</u>": "u"}[tag])
            else:
                if c.content:
                    self._text(out, c.content, fmt)
        return out

    def _text(self, out, text, fmt):
        if self.typography:
            text = _SPACED_HYPHEN.sub(" – ", text)
        pos = 0
        for m in REF_RE.finditer(text):
            if m.start() > pos:
                out.append(("text", text[pos:m.start()], frozenset(fmt)))
            out.append(("ref", m.group(1)))
            pos = m.end()
        if pos < len(text):
            out.append(("text", text[pos:], frozenset(fmt)))


def _find_close(toks, i):
    level = toks[i].level
    j = i + 1
    while j < len(toks):
        if toks[j].level == level and toks[j].nesting == -1:
            return j
        j += 1
    raise MarkdownError("unbalanced markdown tokens")


# ----------------------------------------------------------------------------- post-processing


def _caption_kind(block):
    if not isinstance(block, Paragraph) or not block.inlines or block.inlines[0][0] != "text":
        return None
    m = _CAPTION.match(block.inlines[0][1])
    if not m:
        return None
    word = m.group(1).lower()
    return "table" if word in ("table", "таблица") else "code"


def _postprocess(blocks):
    out = []
    i = 0
    while i < len(blocks):
        b = blocks[i]
        kind = _caption_kind(b)
        if kind:
            inl, attrs = _strip_tail_attrs(_strip_prefix(b.inlines, _CAPTION))
            inl = [x for x in inl if x[0] != "soft"] if inl else inl
            target_cls = Table if kind == "table" else Code
            nxt = blocks[i + 1] if i + 1 < len(blocks) else None
            prev = out[-1] if out else None
            if isinstance(nxt, target_cls) and nxt.caption is None:
                nxt.caption, nxt.id = inl, attrs.get("id") or nxt.id
                _apply_table_attrs(nxt, attrs)
                i += 1
                continue
            if isinstance(prev, target_cls) and prev.caption is None:
                prev.caption, prev.id = inl, attrs.get("id") or prev.id
                _apply_table_attrs(prev, attrs)
                i += 1
                continue
        if isinstance(b, Paragraph) and isinstance(out[-1] if out else None, Math):
            text = plain(b.inlines).lstrip()
            if text.lower().startswith("где"):
                out.append(_where(b.inlines))
                i += 1
                continue
        if isinstance(b, ListBlock):
            for item in b.items:
                item[:] = _postprocess(item)
        out.append(b)
        i += 1
    # mark source lists: ordered list directly under a "sources" heading, or with src ids
    current = None
    for b in out:
        if isinstance(b, Heading) and b.level == 1:
            current = plain(b.inlines).strip().upper()
        if isinstance(b, ListBlock) and b.ordered and (
                current in SOURCES_TITLES or any(x and x.startswith("src:") for x in b.item_ids)):
            b.sources = True
    return out


def _apply_table_attrs(block, attrs):
    if isinstance(block, Table) and attrs.get("widths"):
        try:
            block.widths = [float(x) for x in re.split(r"[,\s]+", attrs["widths"].strip()) if x]
        except ValueError:
            raise MarkdownError(f"bad widths attribute: {attrs['widths']!r}")


def _where(inlines):
    lines, cur = [], []
    for it in inlines:
        if it[0] in ("soft", "br"):
            lines.append(cur)
            cur = []
        else:
            cur.append(it)
    lines.append(cur)
    return Where([ln for ln in lines if ln])
