"""Command line interface: ``gostdoc input.md -o output.docx``."""

from __future__ import annotations

import argparse
import os
import sys
from importlib import resources

import yaml

from . import __version__
from .config import build_config
from .latex2omml import LatexError
from .mdparse import MarkdownError, parse_document
from .numbering import NumberingError
from .render import RenderError, render


def convert(md_path: str, out_path: str, extra_config: dict | None = None, use_word: str = "auto",
            pdf_path: str | None = None, warn=None) -> dict:
    """Convert a Markdown file to .docx. Returns the effective configuration.

    use_word: "auto" (use Word when installed), "yes" (fail without Word), "no".
    """
    from . import wordpost

    with open(md_path, encoding="utf-8") as f:
        src = f.read()
    meta, blocks = parse_document(src, typography=True)
    cfg = build_config(meta, extra_config or {})
    if not cfg["typography"]:  # the setting lives in the front matter, so parse again without it
        meta, blocks = parse_document(src, typography=False)

    if use_word == "yes" and not wordpost.word_available():
        raise RuntimeError("Microsoft Word is not available (needs Windows, Word and pywin32)")
    word = use_word == "yes" or (use_word == "auto" and wordpost.word_available())
    if pdf_path and not word:
        raise RuntimeError("PDF export needs Microsoft Word")

    doc = render(blocks, cfg, os.path.dirname(os.path.abspath(md_path)),
                 update_fields_flag=not word, warn=warn)
    doc.save(out_path)
    if word:
        wordpost.update_with_word(out_path, pdf_path)
    return cfg


def guide() -> str:
    """Return the Markdown syntax guide for authors and AI agents (AI_INSTRUCTIONS.md)."""
    return resources.files(__package__).joinpath("AI_INSTRUCTIONS.md").read_text(encoding="utf-8")


def main(argv=None):
    ap = argparse.ArgumentParser(
        prog="gostdoc",
        description="Convert Markdown to a GOST-formatted .docx (ЕСКД пояснительная записка / ГОСТ 7.32 отчёт).")
    ap.add_argument("input", nargs="?", help="input Markdown file (UTF-8)")
    ap.add_argument("-o", "--output", help="output .docx (default: input name with .docx)")
    ap.add_argument("-c", "--config", help="extra YAML config (overrides front matter)")
    ap.add_argument("--profile", choices=["eskd", "gost732"], help="formatting profile")
    ap.add_argument("--pdf", nargs="?", const="", metavar="PATH",
                    help="also export PDF via Word (default: next to the .docx)")
    ap.add_argument("--word", choices=["auto", "yes", "no"], default="auto",
                    help="use Microsoft Word to update TOC/fields (default: auto)")
    ap.add_argument("--guide", action="store_true",
                    help="print the Markdown syntax guide (written for AI agents) and exit")
    ap.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    args = ap.parse_args(argv)

    if args.guide:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stdout.write(guide())
        return 0
    if not args.input:
        ap.error("the input file is required")

    out = args.output or os.path.splitext(args.input)[0] + ".docx"
    extra = {}
    if args.config:
        with open(args.config, encoding="utf-8") as f:
            extra = yaml.safe_load(f) or {}
    if args.profile:
        extra["profile"] = args.profile
    pdf = None
    if args.pdf is not None:
        pdf = args.pdf or os.path.splitext(out)[0] + ".pdf"
    try:
        convert(args.input, out, extra, use_word=args.word, pdf_path=pdf)
    except (MarkdownError, NumberingError, RenderError, LatexError, RuntimeError, ValueError, OSError) as e:
        print(f"gostdoc: error: {e}", file=sys.stderr)
        return 1
    print(f"written {out}" + (f" and {pdf}" if pdf else ""))
    return 0


if __name__ == "__main__":
    sys.exit(main())
