---
name: gostdoc
description: Write Russian/Belarusian technical documents (пояснительная записка, курсовой/дипломный проект, отчёт по ГОСТ 7.32) as Markdown and convert them to GOST/ЕСКД-formatted .docx with the gostdoc tool. Use when the user wants a document formatted to ГОСТ 2.105, 2.104, 7.32 or ЕСКД, with a frame and title block (рамка, штамп), automatic numbering of figures, tables and formulas, or a GOST table of contents.
---

# gostdoc

1. Make sure the tool is installed: `gostdoc --version`. If it is missing, install it with
   `pip install git+https://github.com/catink123/gostdoc`.
2. Run `gostdoc --guide` and read the whole output before writing anything. It is the
   complete syntax and the rules (numbering, captions, formulas, «где» blocks, sources).
3. Write the document as UTF-8 Markdown, following the guide. Never number headings,
   figures, tables or formulas by hand.
4. Convert: `gostdoc report.md -o report.docx` (add `--pdf` on Windows with Word).
5. Treat every `warning:` line as an error, fix the Markdown and convert again.
6. Run the checklist in section 9 of the guide. If Word was not used, tell the user to open
   the file and accept "Update fields" (or press Ctrl+A, F9) to fill in TOC page numbers.
