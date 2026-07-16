"""Isolated Microsoft Word COM worker used by DOCX previews."""

import os
import sys


def convert_docx_to_pdf(docx_path: str, output_path: str) -> int:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    word = None
    document = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        document = word.Documents.Open(
            os.path.abspath(docx_path),
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
            OpenAndRepair=True,
            NoEncodingDialog=True,
        )
        document.SaveAs(os.path.abspath(output_path), FileFormat=17)
        document.Close(False)
        document = None
        return 0
    except Exception as exc:
        print(str(exc), file=sys.stderr, flush=True)
        return 1
    finally:
        if document is not None:
            try:
                document.Close(False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()