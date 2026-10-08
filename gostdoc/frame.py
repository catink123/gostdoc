"""ESKD sheet frame and title block (основная надпись, ГОСТ 2.104) as header shapes.

Everything is drawn with page-anchored DrawingML shapes, so the body text
flow is not affected. Coordinates are in millimetres from the top-left
corner of an A4 page.

* form 2  (185 x 40 mm) - first sheet of a text document;
* form 2a (185 x 15 mm) - subsequent sheets.
"""

from __future__ import annotations

from xml.sax.saxutils import escape

EMU_PER_MM = 36000
THICK = 25400   # 2 pt  (~0.7 mm, основная линия)
THIN = 12700    # 1 pt  (~0.35 mm, тонкая линия)

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
}
_NSDECL = " ".join(f'xmlns:{k}="{v}"' for k, v in NS.items())

FRAME_X, FRAME_Y, FRAME_W, FRAME_H = 20, 5, 185, 287
BOTTOM = FRAME_Y + FRAME_H  # 292


def _emu(mm: float) -> int:
    return int(round(mm * EMU_PER_MM))


class _Canvas:
    def __init__(self, id_start: int, font: str):
        self.items: list[str] = []
        self.next_id = id_start
        self.font = font

    def _anchor(self, x, y, w, h, graphic, name):
        sid = self.next_id
        self.next_id += 1
        return (
            f'<w:r {_NSDECL}><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" '
            f'relativeHeight="{sid}" behindDoc="0" locked="1" layoutInCell="1" allowOverlap="1">'
            f'<wp:simplePos x="0" y="0"/>'
            f'<wp:positionH relativeFrom="page"><wp:posOffset>{_emu(x)}</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="page"><wp:posOffset>{_emu(y)}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{_emu(w)}" cy="{_emu(h)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'<wp:wrapNone/><wp:docPr id="{sid}" name="{name} {sid}"/><wp:cNvGraphicFramePr/>'
            f'<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
            f'{graphic}</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>'
        )

    def rect(self, x, y, w, h, width=THICK):
        g = (f'<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/>'
             f'<a:ext cx="{_emu(w)}" cy="{_emu(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
             f'<a:noFill/><a:ln w="{width}"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>'
             f'</wps:spPr><wps:bodyPr/></wps:wsp>')
        self.items.append(self._anchor(x, y, w, h, g, "Frame"))

    def hline(self, x1, x2, y, width=THIN):
        self._line(x1, y, x2 - x1, 0, width)

    def vline(self, x, y1, y2, width=THICK):
        self._line(x, y1, 0, y2 - y1, width)

    def _line(self, x, y, w, h, width):
        g = (f'<wps:wsp><wps:cNvCnPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/>'
             f'<a:ext cx="{_emu(w)}" cy="{_emu(h)}"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom>'
             f'<a:ln w="{width}"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>'
             f'</wps:spPr><wps:bodyPr/></wps:wsp>')
        self.items.append(self._anchor(x, y, w, h, g, "Line"))

    def text(self, x, y, w, h, content, size=9, align="center", italic=True, field=None,
             inset=0.5, valign="ctr"):
        """content: str (may contain \\n) or None when ``field`` (raw run XML) is given."""
        rpr = (f'<w:rPr><w:rFonts w:ascii="{self.font}" w:hAnsi="{self.font}" w:cs="{self.font}"/>'
               f'{"<w:i/><w:iCs/>" if italic else ""}<w:sz w:val="{int(size * 2)}"/>'
               f'<w:szCs w:val="{int(size * 2)}"/><w:lang w:val="ru-RU"/></w:rPr>')
        ppr = (f'<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
               f'<w:ind w:left="0" w:right="0" w:firstLine="0"/><w:jc w:val="{align}"/>{rpr}</w:pPr>')
        paras = []
        if field is not None:
            paras.append(f"<w:p>{ppr}{field(rpr)}</w:p>")
        else:
            for line in str(content).split("\n"):
                paras.append(f'<w:p>{ppr}<w:r>{rpr}<w:t xml:space="preserve">{escape(line)}</w:t></w:r></w:p>')
        ins = _emu(inset)
        g = (f'<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/>'
             f'<a:ext cx="{_emu(w)}" cy="{_emu(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
             f'<a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>'
             f'<wps:txbx><w:txbxContent>{"".join(paras)}</w:txbxContent></wps:txbx>'
             f'<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="{ins}" tIns="0" rIns="{ins}" bIns="0" '
             f'anchor="{valign}" anchorCtr="0"><a:noAutofit/></wps:bodyPr></wps:wsp>')
        self.items.append(self._anchor(x, y, w, h, g, "Text"))


def _field_runs(instr_parts):
    """Build a (possibly nested) complex field. instr_parts: list of str | list (nested)."""
    def build(parts, rpr, result):
        out = [f'<w:r>{rpr}<w:fldChar w:fldCharType="begin"/></w:r>']
        for p in parts:
            if isinstance(p, list):
                out.append(build(p, rpr, "1"))
            else:
                out.append(f'<w:r>{rpr}<w:instrText xml:space="preserve">{escape(p)}</w:instrText></w:r>')
        out.append(f'<w:r>{rpr}<w:fldChar w:fldCharType="separate"/></w:r>')
        out.append(f'<w:r>{rpr}<w:t>{result}</w:t></w:r>')
        out.append(f'<w:r>{rpr}<w:fldChar w:fldCharType="end"/></w:r>')
        return "".join(out)
    return lambda rpr: build(instr_parts, rpr, "1")


PAGE_FIELD = _field_runs([" PAGE "])


def _change_block(c: _Canvas, top: float):
    """Left part shared by forms 2 and 2a: two change rows + labels row (heights 5 mm)."""
    xs = [20, 27, 37, 60, 75, 85]
    labels = ["Изм.", "Лист", "№ докум.", "Подпись", "Дата"]
    c.hline(20, 85, top + 5, THIN)
    c.hline(20, 85, top + 10, THICK)
    for k, lab in enumerate(labels):
        c.text(xs[k], top + 10, xs[k + 1] - xs[k], 5, lab, size=9 if k != 2 else 8.5)


def form_2a(cfg, id_start=5000) -> list:
    st = cfg["stamp"]
    c = _Canvas(id_start, cfg["frame_font"])
    c.rect(FRAME_X, FRAME_Y, FRAME_W, FRAME_H, THICK)
    top = BOTTOM - 15
    c.hline(20, 205, top, THICK)
    for x in (27, 37, 60, 75, 85, 195):
        c.vline(x, top, BOTTOM, THICK)
    _change_block(c, top)
    c.hline(195, 205, top + 7, THIN)
    c.text(195, top, 10, 7, "Лист", size=9)
    c.text(195, top + 7, 10, 8, None, size=11, field=PAGE_FIELD)
    c.text(85, top, 110, 15, st.get("designation", ""), size=18)
    return c.items


def form_2(cfg, id_start=4000) -> list:
    st = cfg["stamp"]
    c = _Canvas(id_start, cfg["frame_font"])
    c.rect(FRAME_X, FRAME_Y, FRAME_W, FRAME_H, THICK)
    top = BOTTOM - 40
    c.hline(20, 205, top, THICK)
    # left block verticals: x=27 only spans the change rows, the rest go to the bottom
    c.vline(27, top, top + 15, THICK)
    for x in (37, 60, 75, 85):
        c.vline(x, top, BOTTOM, THICK)
    _change_block(c, top)
    c.hline(20, 205, top + 15, THICK)
    for k in range(1, 5):
        c.hline(20, 85, top + 15 + 5 * k, THIN)
    for k, (role, name) in enumerate(cfg["stamp"]["roles"]):
        y = top + 15 + 5 * k
        c.text(20, y, 17, 5, role, size=9, align="left")
        c.text(37, y, 23, 5, name, size=8 if len(name) > 14 else 9, align="left")
    # designation
    c.text(85, top, 120, 15, st.get("designation", ""), size=18)
    # name + litera block
    c.vline(155, top + 15, BOTTOM, THICK)
    c.hline(155, 205, top + 20, THICK)
    c.hline(155, 205, top + 25, THICK)
    c.vline(170, top + 15, top + 25, THICK)
    c.vline(185, top + 15, top + 25, THICK)
    c.vline(160, top + 20, top + 25, THIN)
    c.vline(165, top + 20, top + 25, THIN)
    c.text(85, top + 15, 70, 25, st.get("title", ""), size=11 if len(st.get("title", "")) > 60 else 12,
           inset=1.5)
    c.text(155, top + 15, 15, 5, "Лит.", size=9)
    c.text(170, top + 15, 15, 5, "Лист", size=9)
    c.text(185, top + 15, 20, 5, "Листов", size=9)
    litera = str(st.get("litera") or "")
    for k, ch in enumerate(litera[:3]):
        c.text(155 + 5 * k, top + 20, 5, 5, ch, size=9)
    c.text(170, top + 20, 15, 5, None, size=9, field=PAGE_FIELD)
    sheets = st.get("sheets")
    if sheets:
        c.text(185, top + 20, 20, 5, str(sheets), size=9)
    else:
        offset = int(cfg["first_page_number"]) - 1
        parts = [" = ", [" NUMPAGES "], f" + {offset} "] if offset else [" NUMPAGES "]
        c.text(185, top + 20, 20, 5, None, size=9, field=_field_runs(parts))
    c.text(155, top + 25, 50, 15, st.get("organization", ""), size=9 if len(st.get("organization", "")) < 70 else 8,
           inset=1)
    return c.items
