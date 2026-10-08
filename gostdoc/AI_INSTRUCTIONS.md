# gostdoc – instructions for AI agents

You are writing a Russian/Belarusian technical document (пояснительная записка к курсовому /
дипломному проекту, отчёт) in **Markdown**, and `gostdoc` turns it into a `.docx` formatted
to GOST (ГОСТ 2.105, 2.104, 7.32). Write content and structure only — **never apply
formatting or numbering by hand**. The tool does fonts, indents, the ЕСКД frame and title
block, the table of contents, and the numbering of sections, figures, tables, listings,
formulas and appendices.

Read this whole file before you write a document. Section 9 has a checklist to run before you
hand the file over.

---

## 1. Running the tool

```bash
# one-time setup
pip install git+https://github.com/catink123/gost-doc-gen
gostdoc --guide                                 # prints this guide

# convert
gostdoc report.md -o report.docx                # or: python -m gostdoc report.md -o report.docx
gostdoc report.md -o report.docx --pdf          # also export PDF (needs MS Word)
gostdoc report.md --profile gost732             # frameless report (ГОСТ 7.32)
gostdoc report.md -c settings.yaml              # settings that override the front matter
gostdoc report.md --word no                     # never start MS Word
```

- Input must be UTF-8. Image paths are resolved relative to the Markdown file.
- When Microsoft Word is installed (Windows), `gostdoc` starts it in the background to fill in
  page numbers in the TOC and the sheet count. This takes 20–60 s. Without Word, the file asks
  "update fields?" when it is opened; answering yes (or pressing F9) fills them in.
- Exit code `0` means success. On failure it prints `gostdoc: error: ...` and exits with `1`.
- `warning: unresolved reference @fig:x` means a cross-reference has no target. The output
  shows `??` in its place. **Treat every warning as an error and fix it.**

Python API: `from gostdoc import convert; convert("in.md", "out.docx", {"profile": "eskd"}, use_word="auto")`.
`gostdoc.guide()` returns this guide as a string.

---

## 2. File skeleton

```markdown
---
profile: eskd                 # eskd (frame + title block, default) | gost732 (no frame)
first_page_number: 5          # sheet number of the first page (СОДЕРЖАНИЕ); title page, задание etc. come before it
stamp:
  designation: АБВГ.505900.001 ПЗ
  title: Программно-аппаратный комплекс диспетчеризации контрольно-пропускных пунктов
  organization: Учреждение образования «Государственный университет», гр. 22-ИТ-1
  litera: У
  developer: Иванов И.И.       # Разраб.
  checker: Петров П.П.         # Провер.
  reviewer: Сидоров С.С.       # Реценз.   (or tech_control: -> "Т. контр.")
  norm_control: Кузнецов К.К.  # Н. контр.
  approver: Смирнов С.С.       # Утв.
---

# ВВЕДЕНИЕ

Текст введения...

# Анализ предметной области

## Описание предметной области

Текст...

# ЗАКЛЮЧЕНИЕ

# СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ

1. Источник... {#src:first}

# ПРИЛОЖЕНИЕ А (обязательное) Техническое задание
```

The output starts with the СОДЕРЖАНИЕ (table of contents) page. The title page and the
assignment sheet (задание) are **not** generated; they are made separately and placed before it.

---

## 3. Headings

| Markdown | Result |
|---|---|
| `# Анализ исходных данных` | `1 АНАЛИЗ ИСХОДНЫХ ДАННЫХ`: new page, bold, upper case, with paragraph indent |
| `## Описание предметной области` | `1.1 Описание предметной области` |
| `### Расчёт затрат` | `1.1.1 Расчёт затрат` |
| `# ВВЕДЕНИЕ` (structural element) | unnumbered, centered, new page |
| `## Что-то {-}` | unnumbered heading |
| `# ПРИЛОЖЕНИЕ А (обязательное) Название` | appendix heading on three lines, centered |
| `# ПРИЛОЖЕНИЕ (справочное) Название` | letter assigned automatically (А, Б, В, Г, Д, Е, Ж, И, К…) |
| `## Раздел` inside an appendix | `А.1 Раздел` (not shown in the TOC) |

Rules:
- **Never type numbers** in headings: write `## Описание`, not `## 1.1 Описание`.
- Write `#` headings in normal sentence case. The ЕСКД profile converts them to upper case.
- No full stop at the end of a heading. Do not underline headings or put them in bold or italic.
- These `#` titles are recognised as unnumbered structural elements: ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ,
  РЕФЕРАТ, АННОТАЦИЯ, СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ, СПИСОК ЛИТЕРАТУРЫ,
  ТЕРМИНЫ И ОПРЕДЕЛЕНИЯ, ПЕРЕЧЕНЬ СОКРАЩЕНИЙ И ОБОЗНАЧЕНИЙ, and others. Write them in upper case.
- Every chapter (`#`) must come before the first appendix.
- Do not put two headings directly one after another without text between them, except
  `#` followed by its first `##`.
- Do not write the СОДЕРЖАНИЕ yourself. It is generated (`toc: false` turns it off).

Add `{#sec:name}` (sections) or `{#app:name}` (appendices) at the end of a heading so you can
refer to it: `## Структура приложения {#sec:struct}`.

---

## 4. Body text

- One paragraph is one block of text with an empty line before and after it. Line breaks
  inside a paragraph are joined with spaces.
- `**bold**`, `*italic*`, `` `code` `` (Courier New), `~~strike~~`, `<sub>x</sub>`, `<sup>2</sup>`.
- Typography is automatic: `"текст"` becomes `«текст»`, `--` becomes `–`, and ` - ` between words
  becomes ` – `. Use `–` for ranges and dashes.
- Hard line break: end the line with `\` or use `<br>`.
- Page break: `\newpage` on its own line. Chapters already start on a new page, so you do not
  need one there.
- `---` (horizontal rule) is also a page break. Do not use it as a decorative separator.

### Lists (перечисления)

```markdown
Комплекс должен обеспечивать:

- приём видеопотока с IP-камер;
- управление шлагбаумом:
  - открытие;
  - закрытие;
- журналирование событий.

Архитектура состоит из трёх уровней:

1. Центральный сервер.
2. Периферийный узел.
3. Рабочее место оператора.
```

- `-` gives a dash bullet (`–`). `1.` gives `1)`, `2)`; nested numbered lists use `а)`, `б)`.
- Always introduce a list with a sentence that ends in a colon.
- Items that are not sentences start in lower case and end with `;`. The last one ends with `.`.
  Items that are full sentences start with a capital and end with `.`.

---

## 5. Figures, tables, listings, formulas

All of these are numbered within their chapter (`2.3`) or appendix (`А.1`). Give each one a
**label** and refer to it from the text with `@label`. The tool inserts the number.

### Figures (рисунки)

```markdown
Структура комплекса показана на рисунке @fig:arch.

![Функциональная структура приложения](images/arch.png){#fig:arch width=150mm}
```

- Put the image on its own paragraph, with empty lines around it.
- The alt text becomes the caption: `Рисунок 1.1 – Функциональная структура приложения`.
  Do not end it with a full stop.
- Optional `width=`: `150mm`, `15cm`, `80%` (of the text width). If you leave it out, the
  image keeps its natural size and is shrunk to fit the page.
- Supported formats: PNG, JPG, GIF, BMP. **SVG is not supported**; convert it to PNG first
  (at least 150 dpi).
- Reference the figure in the text **before** it appears: `на рисунке @fig:arch`, `(рисунок @fig:arch)`.

### Tables (таблицы)

```markdown
Затраты времени приведены в таблице @tbl:time.

Table: Временные затраты на разработку {#tbl:time}

| Этап                  | Недели | Доля, % |
|-----------------------|:------:|:-------:|
| Анализ требований     | 2      | 14,3    |
| Разработка            | 5      | 35,7    |
```

- Put the caption line `Table: Название {#tbl:label}` (or `Таблица: ...`) directly **before**
  the table. It becomes `Таблица 3.1 – Временные затраты на разработку`, aligned left above
  the table.
- The first row is the header. It is centered and repeats automatically on every page the
  table spans.
- Column alignment comes from the separator row: `:---:` center, `---:` right, `---` left.
  Center short numeric columns.
- Column widths are chosen automatically. To set them yourself, give relative widths:
  `Table: Название {#tbl:x widths="3 1 1"}`.
- Merged cells are not supported. Use `<br>` for a line break inside a cell.
- Header cells start with a capital letter and have no full stop at the end.
- A table without a `Table:` line is drawn without a caption or number. GOST requires captions,
  so always add one.

### Code listings (листинги)

````markdown
Логика проверки токенов представлена в листинге @lst:jwt.

Listing: Фрагмент промежуточного ПО для проверки токенов {#lst:jwt}

```go
func Check(token string) error {
    return nil
}
```
````

- Put the caption line `Listing: ...` (or `Листинг: ...`) directly before the fenced code block.
- Code is set in Courier New 12 pt, one paragraph per line, and indentation is kept.
- Keep listings short (fragments). Long code belongs in an appendix.

### Formulas (формулы)

```markdown
Доля времени этапа вычисляется по формуле @eq:share:

$$
T_p = \frac{W_i}{W_s} \cdot 100,
$$ {#eq:share}

где $T_p$ – процент от общего времени;
$W_i$ – число недель одного этапа;
$W_s$ – общее количество недель разработки.
```

- Display formulas: `$$ ... $$`, LaTeX syntax, followed by `{#eq:label}`. The formula is
  centered and its number `(3.1)` is aligned right.
- A reference `@eq:share` prints as `(3.1)`, with parentheses, as GOST requires. So write
  `по формуле @eq:share`, **not** `по формуле (@eq:share)`.
- Punctuation that belongs after the formula (`,` or `.`) goes **inside** the `$$`.
- **Explanation of symbols (экспликация)**: the paragraph directly after the formula that starts
  with `где` is laid out as a list. Put each symbol on its own line, written as
  `$symbol$ – meaning;`, and end the last one with `.`. Do not leave an empty line between the
  formula and `где`.
- Unnumbered formula: `$$ ... $$ {-}`.
- Inline math: `$x_i$`, `$\alpha = 0{,}5$`.
- Use a decimal comma in numbers: write `0{,}5` in LaTeX (a plain `0,5` also works between digits).
- Supported LaTeX: `\frac \sqrt[n]{} ^ _ \sum \prod \int \lim \max \min \sin \cos \ln \log \exp`,
  `\left( \right)`, Greek letters, `\cdot \times \pm \le \ge \ne \approx \infty \to`,
  `\text{...}`, `\mathrm{}`, `\hat \bar \vec \overline`, `cases`, `pmatrix`/`bmatrix`/`matrix`,
  `aligned` with `&` and `\\`. Cyrillic letters inside formulas are fine: `ЗП_{осн}`.

---

## 6. Sources and references to them

```markdown
# СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ

1. Калбертсон, Р. Быстрое тестирование / Р. Калбертсон, К. Браун. – М.: Вильямс, 2002. – 384 с. {#src:testing}
2. WebRTC API [Электронный ресурс]. – Режим доступа: https://developer.mozilla.org/ru/docs/Web/API/WebRTC_API. – Дата доступа: 12.03.2026. {#src:webrtc}
```

In the text: `как показано в [@src:testing]` gives `[1]`. `[@src:testing; @src:webrtc]` is not
supported; write `[@src:testing, @src:webrtc]` instead.

- An ordered list under СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ is printed as `1 Калбертсон, Р. ...`.
- List sources in the order they are first cited in the text.
- Bibliographic records follow ГОСТ 7.1 / 7.0.100. For online sources:
  `Название [Электронный ресурс]. – Режим доступа: URL. – Дата доступа: ДД.ММ.ГГГГ.`
- Every source must be cited in the text at least once.

---

## 7. Labels and references: reference

| Label (definition) | Where it goes | `@label` renders as |
|---|---|---|
| `{#sec:x}` | end of a heading | `1.2` |
| `{#app:x}` | end of an appendix heading | `А` |
| `{#fig:x}` | after `![...](...)` | `1.1` |
| `{#tbl:x}` | end of the `Table:` line | `1.1` |
| `{#lst:x}` | end of the `Listing:` line | `1.1` |
| `{#eq:x}` | after the closing `$$` | `(1.1)` |
| `{#src:x}` | end of a source list item | `1` |

Label names may contain letters, digits, `-`, `_` and `.`. They must be unique.
Phrasing in the text: `на рисунке @fig:x`, `в таблице @tbl:x`, `в листинге @lst:x`,
`по формуле @eq:x`, `в разделе @sec:x`, `в подразделе @sec:y`, `в приложении @app:x`, `[@src:x]`.

---

## 8. Settings (front matter or `-c settings.yaml`)

| Key | Default (eskd) | Meaning |
|---|---|---|
| `profile` | `eskd` | `eskd` = frame + title block, single spacing; `gost732` = no frame, 1.5 spacing, page number at bottom center |
| `first_page_number` | `1` | page/sheet number of the first generated page |
| `toc` / `toc_title` / `toc_levels` | `true` / `СОДЕРЖАНИЕ` / `2` | table of contents |
| `font` / `font_size` | `Times New Roman` / `14` | body font |
| `line_spacing` | `1.0` (`1.5` in gost732) | line spacing multiplier |
| `first_line_indent` | `15` mm | paragraph indent |
| `margins` | `{left: 28, right: 10, top: 15, bottom: 32, header: 9, footer: 20}` | mm |
| `table_font_size` | `14` (`12` in gost732) | font size in tables |
| `code_font` / `code_font_size` | `Courier New` / `12` | listings |
| `math_italic` | `false` | `true` = italic Latin letters in formulas |
| `numbering` | `chapter` | `chapter` → `2.3`, `global` → `7` |
| `number_all_equations` | `true` | `false` = number only formulas with a label |
| `h1_uppercase` | `true` (`false` in gost732) | chapter headings in upper case |
| `typography` | `true` | smart quotes and dashes |
| `frame_font` | `GOST type B` | font of the title block |
| `stamp.sheets` | auto | total number of sheets; by default Word computes `NUMPAGES + first_page_number − 1` |
| `stamp.roles` | see section 2 | full control: `[[Разраб., Иванов И.И.], [Провер., ...], ...]` (5 rows) |

---

## 9. Checklist before you finish

1. `gostdoc file.md -o file.docx` exits with `0` and prints **no warnings**.
2. No hand-made numbers anywhere: no `1.1` in headings, no `Рисунок 2.3` in captions,
   no `(4.1)` after formulas, no `Таблица 1` lines.
3. Every figure, table, listing and formula has a label and is referenced in the text
   **before** it appears.
4. Every formula with symbols is followed by its `где ...` block, one symbol per line.
5. Every source in the list is cited as `[@src:x]`.
6. Headings have no full stop at the end. Captions have no full stop at the end.
7. Lists are introduced with a colon. Items end with `;` and the last one with `.`.
8. Quotes `«»`, dashes `–`, decimal comma `14,3`. Units go after a space: `5 с`, `100 %`.
9. No raw HTML other than `<br>`, `<sub>`, `<sup>`. No SVG images. No merged table cells.
10. If Word was not available, tell the user to open the file and accept "update fields"
    (or press Ctrl+A, F9) so the TOC page numbers appear.

## 10. Not supported (do not try)

Title page and assignment sheet; footnotes; merged table cells; images side by side;
automatic «Продолжение таблицы» labels (the header row repeats on each page instead);
landscape pages; additional columns on the left of the frame (Инв. № подл. etc.).
If you need any of these, generate the document and tell the user which parts to finish
by hand in Word.
