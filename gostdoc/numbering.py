"""Assign GOST numbers to headings, figures, tables, listings, equations and sources."""

from __future__ import annotations

from .mdparse import Code, Figure, Heading, ListBlock, Math, Table

# ГОСТ 2.105 / 7.32: appendix letters skip Ё, З, Й, О, Ч, Ъ, Ы, Ь.
APPENDIX_LETTERS = "АБВГДЕЖИКЛМНПРСТУФХЦШЩЭЮЯ"


class NumberingError(ValueError):
    pass


def assign_numbers(blocks, cfg) -> dict:
    """Mutates blocks (sets .number) and returns label -> number text."""
    labels: dict = {}
    chapter_mode = cfg["numbering"] == "chapter"
    chapter = 0
    sec = [0, 0, 0, 0, 0, 0]
    prefix = None            # "3" in chapter 3, "А" in appendix А, None in unnumbered sections
    in_appendix = False
    counters = {}
    used_letters = set()
    next_letter = 0

    def add_label(key, value):
        if not key:
            return
        if key in labels:
            raise NumberingError(f"duplicate label {key!r}")
        labels[key] = value

    def obj_number(kind):
        counters[kind] = counters.get(kind, 0) + 1
        n = counters[kind]
        if in_appendix or chapter_mode:
            return f"{prefix}.{n}" if prefix else str(n)
        return str(n)

    for b in _walk(blocks):
        if isinstance(b, Heading):
            if b.level == 1:
                sec = [0, 0, 0, 0, 0, 0]
                if b.appendix:
                    in_appendix = True
                    letter = b.appendix_letter
                    if not letter:
                        while APPENDIX_LETTERS[next_letter] in used_letters:
                            next_letter += 1
                        letter = APPENDIX_LETTERS[next_letter]
                    else:
                        if letter in APPENDIX_LETTERS:
                            next_letter = max(next_letter, APPENDIX_LETTERS.index(letter) + 1)
                    used_letters.add(letter)
                    b.appendix_letter = letter
                    b.number = letter
                    prefix = letter
                    counters = {}
                elif b.unnumbered:
                    prefix = None
                    if chapter_mode:
                        counters = {}
                else:
                    if in_appendix:
                        raise NumberingError("numbered chapter after an appendix: "
                                             "chapters must precede appendices")
                    chapter += 1
                    b.number = str(chapter)
                    prefix = str(chapter)
                    if chapter_mode:
                        counters = {}
            else:
                if b.unnumbered or prefix is None:
                    b.unnumbered = True
                else:
                    lvl = b.level - 1
                    sec[lvl] += 1
                    for k in range(lvl + 1, len(sec)):
                        sec[k] = 0
                    parts = [str(x) for x in sec[1:lvl + 1]]
                    b.number = ".".join([prefix] + parts)
            add_label(b.id, b.number or "")
        elif isinstance(b, Figure):
            b.number = obj_number("fig")
            add_label(b.id, b.number)
        elif isinstance(b, Table):
            if b.caption is not None:
                b.number = obj_number("tbl")
            add_label(b.id, b.number)
        elif isinstance(b, Code):
            if b.caption is not None:
                b.number = obj_number("lst")
            add_label(b.id, b.number)
        elif isinstance(b, Math):
            if b.numbered and (cfg["number_all_equations"] or b.id):
                b.number = obj_number("eq")
            add_label(b.id, b.number)
        elif isinstance(b, ListBlock) and b.sources:
            for k, item_id in enumerate(b.item_ids):
                add_label(item_id, str(b.start + k))
    return labels


def _walk(blocks):
    for b in blocks:
        yield b
        if isinstance(b, ListBlock):
            for item in b.items:
                yield from _walk(item)
