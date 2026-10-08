"""gostdoc: Markdown -> GOST-formatted Word documents."""

__version__ = "0.1.0"

from .cli import convert, guide  # noqa: E402

__all__ = ["convert", "guide"]
