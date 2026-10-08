import os
import textwrap

import docx
import pytest
from lxml import etree

from gostdoc import convert
from gostdoc.latex2omml import LatexError, omath
from gostdoc.mdparse import Code, Figure, Heading, Math, Table, Where, parse_document
from gostdoc.numbering import assign_numbers
from gostdoc.config import build_config

HERE = os.path.dirname(__file__)
EXAMPLE = os.path.join(HERE, "..", "examples", "example.md")


@pytest.mark.parametrize("latex", [
    r"T_p = \frac{W_i}{W_s} \cdot 100,",
    r"x = \frac{-b \pm \sqrt{b^2-4ac}}{2a}",
    r"\sum_{i=1}^{n} x_i^2 + \int_0^\infty e^{-x}\,dx",
    r"ЗП_{СР} = 1{,}5 \cdot 14,3",
    r"\left( \frac{a}{b} \right)^2",
    r"f(x) = \begin{cases} 1, & x > 0 \\ 0, & x \le 0 \end{cases}",
    r"\lim_{x \to 0} \frac{\sin x}{x} = 1",
    r"\hat{y} = \overline{x} + \vec{v}",
    r"A = \begin{pmatrix} 1 & 2 \\ 3 & 4 \end{pmatrix}",
    r"a &= b \\ c &= d",
    r"\sqrt[3]{x} + \text{Цена}",
])
def test_latex_produces_wellformed_omml(latex):
    etree.fromstring(omath(latex))


def test_latex_errors_are_reported():
    with pytest.raises(LatexError):
        omath(r"\frac{a}{")


def _numbered(md):
    meta, blocks = parse_document(textwrap.dedent(md))
    labels = assign_numbers(blocks, build_config(meta))
    return blocks, labels


def test_numbering_per_chapter_and_appendix():
    blocks, labels = _numbered("""
        # ВВЕДЕНИЕ

        # Первая глава

        ## Раздел {#sec:a}

        ![Схема](x.png){#fig:a}

        Table: Данные {#tbl:a}

        | a | b |
        |---|---|
        | 1 | 2 |

        $$
        x = 1
        $$ {#eq:a}

        где $x$ – переменная.

        Listing: Код {#lst:a}

        ```
        x
        ```

        # ПРИЛОЖЕНИЕ (обязательное) Расчёты {#app:a}

        ![Ещё](y.png){#fig:b}
        """)
    assert labels == {"sec:a": "1.1", "fig:a": "1.1", "tbl:a": "1.1", "eq:a": "1.1", "lst:a": "1.1",
                      "app:a": "А", "fig:b": "А.1"}
    heads = [b for b in blocks if isinstance(b, Heading)]
    assert heads[0].unnumbered and heads[1].number == "1"
    assert heads[-1].appendix and heads[-1].appendix_status == "обязательное"
    assert any(isinstance(b, Where) for b in blocks)
    assert [type(b) for b in blocks if isinstance(b, (Figure, Table, Code, Math))] == \
        [Figure, Table, Math, Code, Figure]


def test_end_to_end_without_word(tmp_path):
    out = tmp_path / "example.docx"
    convert(EXAMPLE, str(out), use_word="no", warn=lambda m: pytest.fail(m))
    d = docx.Document(str(out))
    texts = [p.text for p in d.paragraphs]
    assert "СОДЕРЖАНИЕ" in texts
    assert "1 АНАЛИЗ ИСХОДНЫХ ДАННЫХ И ПОСТАНОВКА ЗАДАЧИ" in texts
    assert "Рисунок 1.1 – Функциональная структура приложения" in texts
    assert "Таблица 3.1 – Временные затраты на разработку программного обеспечения" in texts
    assert any(t.startswith("Листинг 2.1 – ") for t in texts)
    assert any("по формуле (3.1)" in t for t in texts)
    hdr = d.sections[0].first_page_header._element.xml
    assert "АБВГ.505900.001 ПЗ" in hdr and "Иванов И.И." in hdr
    assert d.sections[0].left_margin.mm == pytest.approx(28, abs=0.1)


def test_short_labels_and_typographed_attributes():
    blocks, labels = _numbered("""
        # Глава

        См. [@src:a, @src:b] и @eq:z.

        Table: Широкая {#tbl:w widths="3 1 1"}

        | a | b | c |
        |---|---|---|
        | 1 | 2 | 3 |

        $$
        z = 0,5
        $$ {#eq:z}

        # СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ

        1. Первый. {#src:a}
        2. Второй. {#src:b}
        """)
    para = blocks[1]
    assert [i[1] for i in para.inlines if i[0] == "ref"] == ["src:a", "src:b", "eq:z"]
    table = next(b for b in blocks if isinstance(b, Table))
    assert table.widths == [3.0, 1.0, 1.0]
    assert labels["src:b"] == "2" and labels["eq:z"] == "1.1"


def test_guide_is_packaged():
    from gostdoc import guide
    assert guide().startswith("# gostdoc")
