"""Render the block IR into a GOST-formatted .docx using python-docx."""

from __future__ import annotations

import os
import re
import sys
from xml.sax.saxutils import escape

from docx import Document
from docx.enum.text import WD_BREAK
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Mm, Pt
from docx.styles.style import StyleFactory
from PIL import Image

from . import frame
from .latex2omml import LatexError, omath, omath_para
from .mdparse import (Code, Figure, Heading, ListBlock, Math, PageBreak, Paragraph, Table, Where,
                      plain)
from .numbering import assign_numbers

TW_PER_MM = 1440 / 25.4


def tw(mm: float) -> int:
    return int(round(mm * TW_PER_MM))


class RenderError(ValueError):
    pass


# ============================================================================ styles

def _style_xml(style_id, name, ppr="", rpr="", based_on="Normal", next_style=None, custom=True,
               stype="paragraph", extra=""):
    custom_attr = ' w:customStyle="1"' if custom else ""
    based = '<w:basedOn w:val="%s"/>' % based_on if based_on else ""
    nxt = '<w:next w:val="%s"/>' % next_style if next_style else ""
    return (
        f'<w:style {nsdecls("w")} w:type="{stype}" w:styleId="{style_id}"{custom_attr}>'
        f'<w:name w:val="{name}"/>{based}{nxt}'
        f'<w:qFormat/>{extra}<w:pPr>{ppr}</w:pPr><w:rPr>{rpr}</w:rPr></w:style>'
    )


def _fonts(font):
    return (f'<w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}" w:cs="{font}"/>')


class Styles:
    NORMAL = "Normal"
    H = {1: "Heading1", 2: "Heading2", 3: "Heading3", 4: "Heading4"}
    APP_H = "AppendixHeading"
    TOC = {1: "TOC1", 2: "TOC2", 3: "TOC3"}
    PLAIN = "NoIndent"
    FIGURE = "FigureImage"
    FIG_CAPTION = "FigureCaption"
    TBL_CAPTION = "TableCaption"
    TBL_TEXT = "TableText"
    CODE = "CodeListing"
    FORMULA = "Formula"
    WHERE = "WhereText"
    LIST = "ListText"


def setup_styles(doc, cfg):
    st = doc.styles.element
    size = int(cfg["font_size"] * 2)
    line = int(round(240 * cfg["line_spacing"]))
    indent = tw(cfg["first_line_indent"])
    font = cfg["font"]
    text_w = tw(cfg["page"]["width"] - cfg["margins"]["left"] - cfg["margins"]["right"])

    # document defaults
    dd = st.find(qn("w:docDefaults"))
    if dd is not None:
        st.remove(dd)
    st.insert(0, parse_xml(
        f'<w:docDefaults {nsdecls("w")}><w:rPrDefault><w:rPr>{_fonts(font)}<w:sz w:val="{size}"/>'
        f'<w:szCs w:val="{size}"/><w:lang w:val="ru-RU" w:eastAsia="ru-RU" w:bidi="ar-SA"/></w:rPr>'
        f'</w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'
        f'</w:pPr></w:pPrDefault></w:docDefaults>'))

    spacing = f'<w:spacing w:before="0" w:after="0" w:line="{line}" w:lineRule="auto"/>'
    base_rpr = f'{_fonts(font)}<w:color w:val="000000"/><w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
    defs = [
        _style_xml("Normal", "Normal", based_on=None, custom=False,
                   ppr=f'<w:widowControl/>{spacing}<w:ind w:firstLine="{indent}"/><w:jc w:val="both"/>',
                   rpr=base_rpr, extra=""),
        _style_xml(Styles.PLAIN, "Без отступа", ppr='<w:ind w:firstLine="0"/>'),
    ]
    for lvl in (1, 2, 3, 4):
        ppr = (f'<w:keepNext/><w:keepLines/>{"<w:pageBreakBefore/>" if lvl == 1 and cfg["h1_page_break"] else ""}'
               f'<w:ind w:firstLine="{indent}"/><w:jc w:val="left"/><w:outlineLvl w:val="{lvl - 1}"/>')
        defs.append(_style_xml(Styles.H[lvl], f"heading {lvl}", ppr=ppr, rpr="<w:b/><w:bCs/>",
                               next_style="Normal", custom=False))
    defs.append(_style_xml(Styles.APP_H, "Заголовок раздела приложения",
                           ppr=f'<w:keepNext/><w:keepLines/><w:ind w:firstLine="{indent}"/><w:jc w:val="left"/>',
                           rpr="<w:b/><w:bCs/>", next_style="Normal"))
    for lvl in (1, 2, 3):
        left = 0 if lvl == 1 else tw(5) * (lvl - 1)
        defs.append(_style_xml(
            Styles.TOC[lvl], f"toc {lvl}", custom=False, next_style="Normal",
            ppr=(f'<w:tabs><w:tab w:val="right" w:leader="dot" w:pos="{text_w}"/></w:tabs>'
                 f'<w:ind w:left="{left}" w:right="567" w:firstLine="0"/><w:jc w:val="left"/>'),
            rpr="<w:noProof/>"))
    defs += [
        _style_xml(Styles.FIGURE, "Рисунок",
                   ppr='<w:keepNext/><w:spacing w:before="120" w:after="0"/><w:ind w:firstLine="0"/>'
                       '<w:jc w:val="center"/>'),
        _style_xml(Styles.FIG_CAPTION, "Подпись рисунка",
                   ppr='<w:spacing w:before="0" w:after="120"/><w:ind w:firstLine="0"/><w:jc w:val="center"/>'),
        _style_xml(Styles.TBL_CAPTION, "Название таблицы",
                   ppr='<w:keepNext/><w:spacing w:before="120" w:after="0"/><w:ind w:firstLine="0"/>'
                       '<w:jc w:val="left"/>'),
        _style_xml(Styles.TBL_TEXT, "Текст таблицы",
                   ppr='<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
                       '<w:ind w:firstLine="0"/><w:jc w:val="left"/>',
                   rpr=f'<w:sz w:val="{int(cfg["table_font_size"] * 2)}"/>'
                       f'<w:szCs w:val="{int(cfg["table_font_size"] * 2)}"/>'),
        _style_xml(Styles.CODE, "Листинг",
                   ppr='<w:spacing w:before="0" w:after="120" w:line="240" w:lineRule="auto"/>'
                       '<w:ind w:firstLine="0"/><w:contextualSpacing/><w:jc w:val="left"/>',
                   rpr=f'{_fonts(cfg["code_font"])}<w:sz w:val="{int(cfg["code_font_size"] * 2)}"/>'
                       f'<w:szCs w:val="{int(cfg["code_font_size"] * 2)}"/><w:lang w:val="en-US"/>'),
        _style_xml(Styles.FORMULA, "Формула",
                   ppr='<w:spacing w:before="120" w:after="120"/><w:ind w:firstLine="0"/><w:jc w:val="center"/>'),
        _style_xml(Styles.WHERE, "Где",
                   ppr='<w:tabs><w:tab w:val="left" w:pos="567"/></w:tabs><w:ind w:left="567" w:hanging="567"/>'
                       '<w:jc w:val="left"/>'),
        _style_xml(Styles.LIST, "Перечисление", ppr=f'<w:ind w:left="0" w:firstLine="{indent}"/>'),
    ]
    for xml in defs:
        el = parse_xml(xml)
        sid = el.get(qn("w:styleId"))
        name = el.find(qn("w:name")).get(qn("w:val"))
        for old in list(st.findall(qn("w:style"))):
            oname = old.find(qn("w:name"))
            if old.get(qn("w:styleId")) == sid or (oname is not None and oname.get(qn("w:val")) == name):
                st.remove(old)
        st.append(el)


# ============================================================================ numbering (lists)

class Lists:
    BULLET, ORDERED, SOURCES = 900, 901, 902

    def __init__(self, doc, cfg):
        self.root = doc.part.numbering_part.element
        self.next_num = 900
        indent = tw(cfg["first_line_indent"])
        step = tw(10)
        bullet = escape(cfg["bullet"])

        def lvl(i, fmt, text, font=None):
            rfonts = f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}" w:hint="default"/>' if font else ""
            return (f'<w:lvl w:ilvl="{i}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/>'
                    f'<w:suff w:val="space"/><w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/>'
                    f'<w:pPr><w:ind w:left="{i * step}" w:firstLine="{indent}"/></w:pPr>'
                    f'<w:rPr>{rfonts}<w:b w:val="0"/><w:i w:val="0"/></w:rPr></w:lvl>')

        bullets = "".join(lvl(i, "bullet", bullet, cfg["font"]) for i in range(9))
        # ГОСТ 2.105: 1) ... а) ... 1) for deeper levels
        ordered_fmts = ["decimal", "russianLower", "decimal"]
        ordered = "".join(lvl(i, ordered_fmts[i % 3], f"%{i + 1})") for i in range(9))
        sources = "".join(lvl(i, "decimal", f"%{i + 1}") for i in range(9))
        abstracts = [(self.BULLET, bullets), (self.ORDERED, ordered), (self.SOURCES, sources)]
        first_num = self.root.find(qn("w:num"))
        for aid, body in abstracts:
            el = parse_xml(f'<w:abstractNum {nsdecls("w")} w:abstractNumId="{aid}">'
                           f'<w:multiLevelType w:val="hybridMultilevel"/>{body}</w:abstractNum>')
            if first_num is not None:
                first_num.addprevious(el)
            else:
                self.root.append(el)

    def new(self, kind, start=1) -> int:
        """Create a numbering instance that restarts at ``start``."""
        self.next_num += 1
        nid = self.next_num
        overrides = "".join(
            f'<w:lvlOverride w:ilvl="{i}"><w:startOverride w:val="{start if i == 0 else 1}"/></w:lvlOverride>'
            for i in range(9))
        self.root.append(parse_xml(f'<w:num {nsdecls("w")} w:numId="{nid}"><w:abstractNumId w:val="{kind}"/>'
                                   f'{overrides}</w:num>'))
        return nid


# ============================================================================ renderer

class Renderer:
    def __init__(self, cfg, base_dir, warn=None):
        self.cfg = cfg
        self.base_dir = base_dir
        self.warn = warn or (lambda m: print("warning: " + m, file=sys.stderr))
        self.doc = Document()
        self.labels = {}
        self.text_w_mm = cfg["page"]["width"] - cfg["margins"]["left"] - cfg["margins"]["right"]
        self.text_h_mm = cfg["page"]["height"] - cfg["margins"]["top"] - cfg["margins"]["bottom"]
        self.has_content = False
        self.after_table = False
        self.in_appendix = False
        self._style_cache = {}

    def _st(self, style_id):
        style = self._style_cache.get(style_id)
        if style is None:
            el = self.doc.styles.element.get_by_id(style_id)
            if el is None:
                raise RenderError(f"internal: style {style_id} missing")
            style = self._style_cache[style_id] = StyleFactory(el)
        return style

    # ------------------------------------------------------------------ public
    def render(self, blocks, update_fields_flag=True):
        doc = self.doc
        # drop the empty paragraph python-docx starts with
        body = doc.element.body
        for p in list(body.findall(qn("w:p"))):
            body.remove(p)
        setup_styles(doc, self.cfg)
        self.lists = Lists(doc, self.cfg)
        self.labels = assign_numbers(blocks, self.cfg)
        self._setup_section()
        self._mark_appendix_headings(blocks)
        if self.cfg["toc"]:
            self._toc(blocks)
        self._blocks(blocks)
        if update_fields_flag:
            settings = doc.settings.element
            settings.append(parse_xml(f'<w:updateFields {nsdecls("w")} w:val="true"/>'))
        self._core_props()
        return doc

    # ------------------------------------------------------------------ page setup
    def _setup_section(self):
        cfg, m = self.cfg, self.cfg["margins"]
        sec = self.doc.sections[0]
        sec.page_width, sec.page_height = Mm(cfg["page"]["width"]), Mm(cfg["page"]["height"])
        sec.left_margin, sec.right_margin = Mm(m["left"]), Mm(m["right"])
        sec.top_margin, sec.bottom_margin = Mm(m["top"]), Mm(m["bottom"])
        sec.header_distance, sec.footer_distance = Mm(m["header"]), Mm(m["footer"])
        sect = sec._sectPr
        for old in sect.findall(qn("w:pgNumType")):
            sect.remove(old)
        pg = parse_xml(f'<w:pgNumType {nsdecls("w")} w:start="{int(cfg["first_page_number"])}"/>')
        cols = sect.find(qn("w:cols"))
        if cols is not None:
            cols.addprevious(pg)
        else:
            sect.append(pg)

        if cfg["frame"]:
            sec.different_first_page_header_footer = True
            self._fill_header(sec.first_page_header, frame.form_2(cfg))
            self._fill_header(sec.header, frame.form_2a(cfg))
            # first sheet: push the text body above the 40 mm title block
            need = frame.BOTTOM - 40 - 2  # mm from top where text must end
            spacer = max(0.0, (cfg["page"]["height"] - need) - m["footer"])
            fp = sec.first_page_footer.paragraphs[0]
            fp._p.get_or_add_pPr().append(parse_xml(
                f'<w:spacing {nsdecls("w")} w:before="0" w:after="0" w:line="{tw(spacer)}" w:lineRule="exact"/>'))
            self._set_ind0(fp)
            for p in (sec.footer.paragraphs[0],):
                self._set_ind0(p)
        else:
            p = sec.footer.paragraphs[0]
            self._set_ind0(p)
            p._p.get_or_add_pPr().append(parse_xml(f'<w:jc {nsdecls("w")} w:val="center"/>'))
            self._field(p, " PAGE ", "1")

    def _fill_header(self, header, runs):
        p = header.paragraphs[0]
        self._set_ind0(p)
        p._p.get_or_add_pPr().append(parse_xml(
            f'<w:spacing {nsdecls("w")} w:before="0" w:after="0" w:line="20" w:lineRule="exact"/>'))
        for r in runs:
            p._p.append(parse_xml(r))

    @staticmethod
    def _set_ind0(p):
        p._p.get_or_add_pPr().append(parse_xml(f'<w:ind {nsdecls("w")} w:firstLine="0"/>'))

    def _field(self, p, instr, result):
        p._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:fldChar w:fldCharType="begin"/></w:r>'))
        p._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:instrText xml:space="preserve">{escape(instr)}</w:instrText></w:r>'))
        p._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:fldChar w:fldCharType="separate"/></w:r>'))
        p._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:t>{escape(result)}</w:t></w:r>'))
        p._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:fldChar w:fldCharType="end"/></w:r>'))

    def _core_props(self):
        cp = self.doc.core_properties
        st = self.cfg["stamp"]
        cp.title = str(self.cfg.get("title") or st.get("title") or "")
        roles = st.get("roles") or []
        cp.author = str(self.cfg.get("author") or (roles[0][1] if roles else "") or "")
        cp.language = "ru-RU"

    # ------------------------------------------------------------------ TOC
    def _toc(self, blocks):
        doc = self.doc
        p = doc.add_paragraph(style=self._st(Styles.PLAIN))
        p.alignment = 1
        r = p.add_run(self.cfg["toc_title"])
        r.bold = True
        doc.add_paragraph(style=self._st(Styles.PLAIN))
        levels = int(self.cfg["toc_levels"])
        entries = []
        for b in blocks:
            if isinstance(b, Heading) and b.level <= levels:
                if b.appendix:
                    entries.append((1, self._appendix_toc_text(b)))
                elif not getattr(b, "_in_appendix", False):
                    entries.append((b.level, self._heading_text(b)))
        # The field result is pre-filled with the entries (without page numbers);
        # Word recomputes it on update (gostdoc does that automatically when Word is available).
        first_lvl = min(entries[0][0], 3) if entries else 1
        first = doc.add_paragraph(style=self._st(Styles.TOC[first_lvl]))
        first._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:fldChar w:fldCharType="begin"/></w:r>'))
        first._p.append(parse_xml(
            f'<w:r {nsdecls("w")}><w:instrText xml:space="preserve"> TOC \\o "1-{levels}" \\h \\z \\u </w:instrText></w:r>'))
        first._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:fldChar w:fldCharType="separate"/></w:r>'))
        last = first
        for k, (lvl, text) in enumerate(entries):
            p = first if k == 0 else doc.add_paragraph(style=self._st(Styles.TOC[min(lvl, 3)]))
            p.add_run(text + "\t")
            last = p
        last._p.append(parse_xml(f'<w:r {nsdecls("w")}><w:fldChar w:fldCharType="end"/></w:r>'))
        self.has_content = True

    def _appendix_toc_text(self, h):
        parts = [f"ПРИЛОЖЕНИЕ {h.appendix_letter}"]
        if h.appendix_status:
            parts.append(f"({h.appendix_status})")
        title = plain(h.inlines).strip()
        if title:
            parts.append(title)
        return " ".join(parts)

    def _heading_text(self, h):
        text = self._plain_resolved(h.inlines).strip()
        if h.level == 1 and self.cfg["h1_uppercase"]:
            text = text.upper()
        return f"{h.number} {text}" if h.number and not h.unnumbered else text

    def _plain_resolved(self, inlines):
        out = []
        for it in inlines:
            if it[0] == "ref":
                out.append(self._ref(it[1]))
            elif it[0] in ("text", "math"):
                out.append(it[1])
            else:
                out.append(" ")
        return "".join(out)

    # ------------------------------------------------------------------ blocks
    @staticmethod
    def _mark_appendix_headings(blocks):
        """Subsections of appendices get their own style and stay out of the TOC."""
        in_app = False
        for b in blocks:
            if isinstance(b, Heading) and b.level == 1:
                in_app = b.appendix
            elif isinstance(b, Heading):
                b._in_appendix = in_app

    def _blocks(self, blocks):
        prev = None
        n = len(blocks)
        for i, b in enumerate(blocks):
            nxt = blocks[i + 1] if i + 1 < n else None
            if isinstance(b, Heading):
                self._heading(b, prev, nxt)
            elif isinstance(b, Paragraph):
                self._paragraph(b, keep_next=isinstance(nxt, Math))
            elif isinstance(b, Figure):
                self._figure(b)
            elif isinstance(b, Table):
                self._table(b)
            elif isinstance(b, Code):
                self._code(b)
            elif isinstance(b, Math):
                self._math(b, keep_next=isinstance(nxt, Where))
            elif isinstance(b, Where):
                self._where(b)
            elif isinstance(b, ListBlock):
                self._list(b, 0)
            elif isinstance(b, PageBreak):
                if not (isinstance(nxt, Heading) and nxt.level == 1 and self.cfg["h1_page_break"]):
                    p = self.doc.add_paragraph(style=self._st(Styles.PLAIN))
                    p.add_run().add_break(WD_BREAK.PAGE)
            if not isinstance(b, Table):
                self.after_table = False
            prev = b
            self.has_content = True

    def _spacers(self, count, keep=False):
        for _ in range(count):
            p = self.doc.add_paragraph(style=self._st(Styles.PLAIN))
            if keep:
                p.paragraph_format.keep_with_next = True

    def _heading(self, h, prev, nxt):
        doc = self.doc
        next_is_heading = isinstance(nxt, Heading)
        if h.level == 1:
            self.in_appendix = h.appendix
            p = doc.add_paragraph(style=self._st(Styles.H[1]))
            if not self.has_content:
                p.paragraph_format.page_break_before = False
            if h.appendix:
                self._center(p)
                p.add_run(f"ПРИЛОЖЕНИЕ {h.appendix_letter}")
                if h.appendix_status:
                    p.add_run().add_break()
                    r = p.add_run(f"({h.appendix_status})")
                    r.bold = False
                if h.inlines:
                    p.add_run().add_break()
                    self._inlines(p, h.inlines)
            else:
                if h.unnumbered:
                    self._center(p)
                elif h.number:
                    p.add_run(f"{h.number} ")
                upper = self.cfg["h1_uppercase"]
                self._inlines(p, h.inlines, upper=upper)
            self._spacers(1 if next_is_heading else 2, keep=True)
            return
        if not isinstance(prev, Heading):
            self._spacers(2)
        level = min(h.level, 4)
        style = Styles.APP_H if getattr(h, "_in_appendix", False) else Styles.H[level]
        p = doc.add_paragraph(style=self._st(style))
        if h.number and not h.unnumbered:
            p.add_run(f"{h.number} ")
        self._inlines(p, h.inlines)
        self._spacers(1 if next_is_heading else 2, keep=True)

    @staticmethod
    def _center(p):
        p.alignment = 1
        p.paragraph_format.first_line_indent = 0

    def _paragraph(self, b, keep_next=False):
        p = self.doc.add_paragraph(style=self._st(Styles.NORMAL))
        if keep_next:  # "... по формуле:" stays on the page of its formula
            p.paragraph_format.keep_with_next = True
        if self.after_table:
            p.paragraph_format.space_before = Pt(6)
        self._inlines(p, b.inlines)

    # ------------------------------------------------------------------ lists
    def _list(self, b, level):
        if b.sources:
            kind = Lists.SOURCES
        else:
            kind = Lists.ORDERED if b.ordered else Lists.BULLET
        num_id = self.lists.new(kind, b.start)
        self._list_items(b, num_id, level)

    def _list_items(self, b, num_id, level):
        for item in b.items:
            first = True
            for blk in item:
                if isinstance(blk, Paragraph):
                    p = self.doc.add_paragraph(style=self._st(Styles.LIST))
                    if first:
                        p._p.get_or_add_pPr().append(parse_xml(
                            f'<w:numPr {nsdecls("w")}><w:ilvl w:val="{level}"/><w:numId w:val="{num_id}"/></w:numPr>'))
                    else:
                        p.paragraph_format.left_indent = Mm(10 * level)
                        p.paragraph_format.first_line_indent = Mm(self.cfg["first_line_indent"])
                    self._inlines(p, blk.inlines)
                    first = False
                elif isinstance(blk, ListBlock):
                    if blk.ordered == b.ordered and not blk.sources:
                        self._list_items(blk, num_id, level + 1)
                    else:
                        kind = Lists.ORDERED if blk.ordered else Lists.BULLET
                        self._list_items(blk, self.lists.new(kind, blk.start), level + 1)
                else:
                    self._blocks([blk])
                    first = False

    # ------------------------------------------------------------------ figures
    def _figure(self, b):
        path = b.src if os.path.isabs(b.src) else os.path.join(self.base_dir, b.src)
        if not os.path.exists(path):
            raise RenderError(f"image not found: {b.src}")
        if path.lower().endswith(".svg"):
            raise RenderError(f"SVG is not supported, convert to PNG: {b.src}")
        width_mm = self._image_width(path, b.width)
        p = self.doc.add_paragraph(style=self._st(Styles.FIGURE))
        p.add_run().add_picture(path, width=Mm(width_mm))
        cap = self.doc.add_paragraph(style=self._st(Styles.FIG_CAPTION))
        cap.add_run(f"Рисунок {b.number} {self.cfg['caption_dash']} ")
        self._inlines(cap, b.caption)

    def _image_width(self, path, spec):
        if spec:
            s = spec.strip().lower()
            m = re.fullmatch(r"([\d.]+)\s*(%|mm|cm|px|in)?", s)
            if not m:
                raise RenderError(f"bad image width {spec!r} (use e.g. 120mm, 12cm, 80%)")
            v, unit = float(m.group(1)), m.group(2) or "mm"
            mm = {"%": self.text_w_mm * v / 100, "mm": v, "cm": v * 10, "px": v * 25.4 / 96,
                  "in": v * 25.4}[unit]
            return min(mm, self.text_w_mm)
        with Image.open(path) as im:
            w, h = im.size
            dpi = im.info.get("dpi", (96, 96))[0] or 96
        if dpi < 50:
            dpi = 96
        mm = w * 25.4 / dpi
        mm = min(mm, self.text_w_mm)
        max_h = self.text_h_mm * 0.8
        if mm * h / w > max_h:
            mm = max_h * w / h
        return mm

    # ------------------------------------------------------------------ tables
    def _table(self, b):
        doc = self.doc
        if b.caption is not None:
            cap = doc.add_paragraph(style=self._st(Styles.TBL_CAPTION))
            cap.add_run(f"Таблица {b.number} {self.cfg['caption_dash']} ")
            self._inlines(cap, b.caption)
        rows = b.rows
        ncols = max(len(r) for r in rows)
        total = tw(self.text_w_mm)
        widths = self._col_widths(b, ncols, total)
        t = doc.add_table(rows=len(rows), cols=ncols)
        t.style = self._table_grid()
        tbl = t._tbl
        tblPr = tbl.tblPr
        for tag in ("w:tblW", "w:tblLayout", "w:tblLook"):
            for el in tblPr.findall(qn(tag)):
                tblPr.remove(el)
        tblPr.append(parse_xml(f'<w:tblW {nsdecls("w")} w:w="{sum(widths)}" w:type="dxa"/>'))
        tblPr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))
        grid = tbl.tblGrid
        for gc, w in zip(grid.findall(qn("w:gridCol")), widths):
            gc.set(qn("w:w"), str(w))
        for ri, row in enumerate(rows):
            tr = t.rows[ri]
            if ri < b.header_rows:
                trPr = tr._tr.get_or_add_trPr()
                trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
            trPr = tr._tr.get_or_add_trPr()
            trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
            for ci in range(ncols):
                cell = tr.cells[ci]
                tcPr = cell._tc.get_or_add_tcPr()
                for el in tcPr.findall(qn("w:tcW")):
                    tcPr.remove(el)
                tcPr.append(parse_xml(f'<w:tcW {nsdecls("w")} w:w="{widths[ci]}" w:type="dxa"/>'))
                tcPr.append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>'))
                p = cell.paragraphs[0]
                p.style = self._st(Styles.TBL_TEXT)
                inl = row[ci] if ci < len(row) else []
                if ri < b.header_rows:
                    p.alignment = 1
                else:
                    al = b.aligns[ci] if ci < len(b.aligns) else None
                    p.alignment = {"center": 1, "right": 2}.get(al, 0)
                self._inlines(p, inl)
        self.after_table = True

    def _table_grid(self):
        el = self.doc.styles.element.get_by_id("TableGrid")
        if el is None:
            raise RenderError("internal: Table Grid style missing")
        return StyleFactory(el)

    def _col_widths(self, b, ncols, total):
        if b.widths:
            if len(b.widths) != ncols:
                raise RenderError(f"table {b.number or ''}: widths has {len(b.widths)} values, "
                                  f"table has {ncols} columns")
            weights = b.widths
        else:
            weights = self._auto_widths(b, ncols, total)
        s = sum(weights)
        widths = [int(total * w / s) for w in weights]
        widths[-1] += total - sum(widths)
        return widths

    def _auto_widths(self, b, ncols, total):
        """Estimate column widths (twips) like a browser's auto table layout: no column
        narrower than its longest word, extra space shared by how much text a column has."""
        char_w = tw(self.cfg["table_font_size"] * 0.55 * 25.4 / 72)   # average glyph width
        pad = tw(4.5)                                                  # cell margins + slack
        mins, prefs = [], []
        for ci in range(ncols):
            texts = [plain(r[ci]) if ci < len(r) else "" for r in b.rows]
            longest_word = max((len(w) for t in texts for w in t.split()), default=1)
            longest_cell = max((len(t) for t in texts), default=1)
            mins.append(longest_word * char_w + pad)
            prefs.append(min(longest_cell * char_w + pad, total))
        if sum(prefs) <= total:
            return prefs
        if sum(mins) >= total:
            return mins
        spare = total - sum(mins)
        flex = [p - m for p, m in zip(prefs, mins)]
        return [m + spare * f / sum(flex) for m, f in zip(mins, flex)]

    # ------------------------------------------------------------------ code
    def _code(self, b):
        doc = self.doc
        if b.caption is not None:
            cap = doc.add_paragraph(style=self._st(Styles.TBL_CAPTION))
            cap.add_run(f"Листинг {b.number} {self.cfg['caption_dash']} ")
            self._inlines(cap, b.caption)
        for line in b.lines:
            p = doc.add_paragraph(style=self._st(Styles.CODE))
            text = line.replace("\t", "    ")
            if text:
                p.add_run(text)

    # ------------------------------------------------------------------ math
    def _omml(self, latex, display):
        fn = omath_para if display else omath
        try:
            return parse_xml(fn(latex, italic=self.cfg["math_italic"], text_font=self.cfg["font"]))
        except LatexError as e:
            raise RenderError(f"formula {latex!r}: {e}")

    def _math(self, b, keep_next=False):
        doc = self.doc
        if not b.number:
            p = doc.add_paragraph(style=self._st(Styles.FORMULA))
            p.paragraph_format.keep_with_next = keep_next
            p._p.append(self._omml(b.latex, True))
            return
        num_w = tw(15)
        total = tw(self.text_w_mm)
        widths = [num_w, total - 2 * num_w, num_w]
        t = doc.add_table(rows=1, cols=3)
        tblPr = t._tbl.tblPr
        for tag in ("w:tblW", "w:tblLook"):
            for el in tblPr.findall(qn(tag)):
                tblPr.remove(el)
        tblPr.append(parse_xml(f'<w:tblW {nsdecls("w")} w:w="{total}" w:type="dxa"/>'))
        tblPr.append(parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>'))
        tblPr.append(parse_xml(f'<w:tblCellMar {nsdecls("w")}><w:left w:w="0" w:type="dxa"/>'
                               f'<w:right w:w="0" w:type="dxa"/></w:tblCellMar>'))
        for gc, w in zip(t._tbl.tblGrid.findall(qn("w:gridCol")), widths):
            gc.set(qn("w:w"), str(w))
        row = t.rows[0]
        row._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        for ci, cell in enumerate(row.cells):
            tcPr = cell._tc.get_or_add_tcPr()
            tcPr.append(parse_xml(f'<w:tcW {nsdecls("w")} w:w="{widths[ci]}" w:type="dxa"/>'))
            tcPr.append(parse_xml(f'<w:vAlign {nsdecls("w")} w:val="center"/>'))
            p = cell.paragraphs[0]
            p.style = self._st(Styles.FORMULA)
            p.paragraph_format.keep_with_next = keep_next
        row.cells[1].paragraphs[0]._p.append(self._omml(b.latex, True))
        np_ = row.cells[2].paragraphs[0]
        np_.alignment = 2
        np_.add_run(f"({b.number})")

    def _where(self, b):
        for k, line in enumerate(b.lines):
            p = self.doc.add_paragraph(style=self._st(Styles.WHERE))
            if k == 0:
                line = _strip_where(line)
                p.add_run("где\t")
            else:
                p.add_run("\t")
            self._inlines(p, line)

    # ------------------------------------------------------------------ inlines
    def _ref(self, key):
        if key in self.labels and self.labels[key]:
            # ГОСТ 2.105: references to formulas are given in parentheses
            return f"({self.labels[key]})" if key.startswith("eq:") else self.labels[key]
        self.warn(f"unresolved reference @{key}")
        return "??"

    def _inlines(self, p, inlines, upper=False):
        for it in inlines:
            kind = it[0]
            if kind == "text":
                text = it[1].upper() if upper else it[1]
                r = p.add_run(text)
                fmt = it[2]
                if "b" in fmt:
                    r.bold = True
                if "i" in fmt:
                    r.italic = True
                if "s" in fmt:
                    r.font.strike = True
                if "u" in fmt:
                    r.underline = True
                if "sub" in fmt:
                    r.font.subscript = True
                if "sup" in fmt:
                    r.font.superscript = True
                if "code" in fmt:
                    r.font.name = self.cfg["code_font"]
                    r._r.get_or_add_rPr().get_or_add_rFonts().set(qn("w:cs"), self.cfg["code_font"])
            elif kind == "soft":
                p.add_run(" ")
            elif kind == "br":
                p.add_run().add_break()
            elif kind == "math":
                p._p.append(self._omml(it[1], False))
            elif kind == "ref":
                p.add_run(self._ref(it[1]))


def _strip_where(line):
    if line and line[0][0] == "text":
        t = re.sub(r"^\s*где\s*", "", line[0][1], flags=re.I)
        return ([("text", t, line[0][2])] if t else []) + list(line[1:])
    return line


def render(blocks, cfg, base_dir, update_fields_flag=True, warn=None):
    return Renderer(cfg, base_dir, warn).render(blocks, update_fields_flag)
