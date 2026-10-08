"""Optional post-processing through Microsoft Word (Windows only).

Word is the only engine that can paginate the document, so it is used to
fill in the table of contents page numbers and sheet counts, and optionally
to export a PDF. Without Word the .docx asks to update fields when opened.
"""

from __future__ import annotations

import os
import sys


def word_available() -> bool:
    if sys.platform != "win32":
        return False
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return False
    try:
        import winreg
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"Word.Application\CLSID"))
        return True
    except OSError:
        return False


def update_with_word(docx_path: str, pdf_path: str | None = None) -> None:
    import pythoncom
    import win32com.client

    docx_path = os.path.abspath(docx_path)
    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    try:
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(docx_path, ConfirmConversions=False, ReadOnly=False,
                                  AddToRecentFiles=False, Visible=False)
        try:
            doc.Repaginate()
            for toc in doc.TablesOfContents:
                toc.Update()
            doc.Fields.Update()
            for sec in doc.Sections:
                for coll in (sec.Headers, sec.Footers):
                    for k in (1, 2, 3):
                        try:
                            hf = coll(k)
                            hf.Range.Fields.Update()
                            for shape in hf.Shapes:
                                try:
                                    shape.TextFrame.TextRange.Fields.Update()
                                except Exception:
                                    pass
                        except Exception:
                            pass
            doc.Repaginate()
            for toc in doc.TablesOfContents:
                toc.UpdatePageNumbers()
            doc.Save()
            if pdf_path:
                doc.ExportAsFixedFormat(os.path.abspath(pdf_path), 17)  # wdExportFormatPDF
        finally:
            doc.Close(SaveChanges=0)
    finally:
        word.Quit()
        pythoncom.CoUninitialize()
