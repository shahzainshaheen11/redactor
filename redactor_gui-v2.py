#!/usr/bin/env python3
"""GUI redactor for TXT, DOCX, DOC, and PDF files."""

from __future__ import annotations

import ctypes
from pathlib import Path
import re
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import pymupdf
except ImportError:
    try:
        import fitz as pymupdf  # type: ignore[no-redef]
    except ImportError:
        pymupdf = None

try:
    from docx import Document
except ImportError:
    Document = None

try:
    import win32com.client as win32_client
except ImportError:
    win32_client = None

BOX = "\u2588"
SUPPORTED_SUFFIXES = {".txt", ".docx", ".doc", ".pdf"}
WORD_DOCX_FORMAT = 16


def enable_high_dpi() -> None:
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def default_output_path(input_path: Path) -> Path:
    suffix = ".docx" if input_path.suffix.lower() == ".doc" else input_path.suffix
    return input_path.with_name(f"{input_path.stem}_redacted{suffix}")


def redact_text(text: str, target: str, whole_word: bool, ignore_case: bool) -> tuple[str, int]:
    pattern = re.escape(target)
    if whole_word:
        pattern = rf"\b{pattern}\b"

    flags = re.IGNORECASE if ignore_case else 0
    regex = re.compile(pattern, flags)

    def replacer(match: re.Match[str]) -> str:
        return BOX * len(match.group(0))

    return regex.subn(replacer, text)


def redact_txt_file(
    input_path: Path,
    output_path: Path,
    target: str,
    whole_word: bool,
    ignore_case: bool,
) -> int:
    text = input_path.read_text(encoding="utf-8")
    redacted_text, count = redact_text(text, target, whole_word, ignore_case)
    output_path.write_text(redacted_text, encoding="utf-8")
    return count


def redact_docx_paragraph(paragraph, target: str, whole_word: bool, ignore_case: bool) -> int:
    count = 0
    if paragraph.runs:
        for run in paragraph.runs:
            if not run.text:
                continue
            redacted_text, replacements = redact_text(run.text, target, whole_word, ignore_case)
            if replacements:
                run.text = redacted_text
                count += replacements
        return count

    redacted_text, replacements = redact_text(paragraph.text, target, whole_word, ignore_case)
    if replacements:
        paragraph.text = redacted_text
    return replacements


def redact_docx_container(container, target: str, whole_word: bool, ignore_case: bool) -> int:
    count = 0
    for paragraph in container.paragraphs:
        count += redact_docx_paragraph(paragraph, target, whole_word, ignore_case)

    for table in getattr(container, "tables", []):
        for row in table.rows:
            for cell in row.cells:
                count += redact_docx_container(cell, target, whole_word, ignore_case)

    return count


def redact_docx_file(
    input_path: Path,
    output_path: Path,
    target: str,
    whole_word: bool,
    ignore_case: bool,
) -> int:
    if Document is None:
        raise RuntimeError("DOCX support requires python-docx. Install it with: pip install python-docx")

    document = Document(str(input_path))
    count = redact_docx_container(document, target, whole_word, ignore_case)

    for section in document.sections:
        count += redact_docx_container(section.header, target, whole_word, ignore_case)
        count += redact_docx_container(section.footer, target, whole_word, ignore_case)

    document.save(str(output_path))
    return count


def convert_doc_to_docx(input_path: Path) -> Path:
    if win32_client is None:
        raise RuntimeError(
            "DOC support requires pywin32 and Microsoft Word on Windows. "
            "Install pywin32 with: pip install pywin32"
        )

    temp_file = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
    temp_file.close()
    temp_path = Path(temp_file.name)

    word = None
    document = None
    try:
        word = win32_client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        document = word.Documents.Open(str(input_path))
        document.SaveAs2(str(temp_path), FileFormat=WORD_DOCX_FORMAT)
    except Exception as exc:
        try:
            temp_path.unlink(missing_ok=True)
        except OSError:
            pass
        raise RuntimeError(
            "Could not open this .doc file through Microsoft Word. "
            "Make sure Microsoft Word is installed."
        ) from exc
    finally:
        if document is not None:
            document.Close(False)
        if word is not None:
            word.Quit()

    return temp_path


def redact_doc_file(
    input_path: Path,
    output_path: Path,
    target: str,
    whole_word: bool,
    ignore_case: bool,
) -> int:
    temp_docx = convert_doc_to_docx(input_path)
    try:
        return redact_docx_file(temp_docx, output_path, target, whole_word, ignore_case)
    finally:
        try:
            temp_docx.unlink(missing_ok=True)
        except OSError:
            pass


def redact_pdf_file(
    input_path: Path,
    output_path: Path,
    target: str,
    whole_word: bool,
    ignore_case: bool,
) -> int:
    if pymupdf is None:
        raise RuntimeError("PDF support requires PyMuPDF. Install it with: pip install pymupdf")

    document = pymupdf.open(str(input_path))
    count = 0
    try:
        for page in document:
            page_hits = 0
            for x0, y0, x1, y1, word_text, *_ in page.get_text("words"):
                redacted_word, replacements = redact_text(
                    word_text,
                    target,
                    whole_word,
                    ignore_case,
                )
                if not replacements:
                    continue

                rect = pymupdf.Rect(x0, y0, x1, y1)
                page.add_redact_annot(
                    rect,
                    text=redacted_word,
                    fontname="Helv",
                    fontsize=11,
                    fill=(0, 0, 0),
                    text_color=(1, 1, 1),
                    cross_out=False,
                )
                count += replacements
                page_hits += 1

            if page_hits:
                page.apply_redactions()

        document.save(str(output_path))
    finally:
        document.close()

    return count


def redact_supported_file(
    input_path: Path,
    output_path: Path,
    target: str,
    whole_word: bool,
    ignore_case: bool,
) -> tuple[Path, int]:
    suffix = input_path.suffix.lower()

    if suffix == ".txt":
        return output_path, redact_txt_file(input_path, output_path, target, whole_word, ignore_case)
    if suffix == ".docx":
        return output_path, redact_docx_file(input_path, output_path, target, whole_word, ignore_case)
    if suffix == ".pdf":
        return output_path, redact_pdf_file(input_path, output_path, target, whole_word, ignore_case)
    if suffix == ".doc":
        docx_output = output_path.with_suffix(".docx")
        return docx_output, redact_doc_file(input_path, docx_output, target, whole_word, ignore_case)

    raise RuntimeError("Unsupported file type.")


class RedactorApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("File Redactor")
        self.root.resizable(False, False)
        self.root.columnconfigure(1, weight=1)

        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.target_var = tk.StringVar()
        self.whole_word_var = tk.BooleanVar(value=True)
        self.case_sensitive_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(
            value="Choose a .txt, .pdf, .docx, or .doc file to begin."
        )

        self._build_ui()

    def _build_ui(self) -> None:
        padding = {"padx": 10, "pady": 6}
        frame = ttk.Frame(self.root, padding=14)
        frame.grid(sticky="nsew")
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="File").grid(row=0, column=0, sticky="w", **padding)
        ttk.Entry(frame, textvariable=self.input_var, width=48).grid(
            row=0, column=1, sticky="ew", **padding
        )
        ttk.Button(frame, text="Browse", command=self.browse_input).grid(
            row=0, column=2, sticky="ew", **padding
        )

        ttk.Label(frame, text="Word or phrase").grid(row=1, column=0, sticky="w", **padding)
        ttk.Entry(frame, textvariable=self.target_var, width=48).grid(
            row=1, column=1, columnspan=2, sticky="ew", **padding
        )

        ttk.Label(frame, text="Save as").grid(row=2, column=0, sticky="w", **padding)
        ttk.Entry(frame, textvariable=self.output_var, width=48).grid(
            row=2, column=1, sticky="ew", **padding
        )
        ttk.Button(frame, text="Browse", command=self.browse_output).grid(
            row=2, column=2, sticky="ew", **padding
        )

        ttk.Checkbutton(frame, text="Whole word only", variable=self.whole_word_var).grid(
            row=3, column=0, columnspan=2, sticky="w", **padding
        )
        ttk.Checkbutton(frame, text="Case sensitive", variable=self.case_sensitive_var).grid(
            row=4, column=0, columnspan=2, sticky="w", **padding
        )

        ttk.Button(frame, text="Redact File", command=self.redact_file).grid(
            row=5, column=0, columnspan=3, sticky="ew", padx=10, pady=(10, 6)
        )

        ttk.Label(
            frame, textvariable=self.status_var, wraplength=440, justify="left"
        ).grid(row=6, column=0, columnspan=3, sticky="w", padx=10, pady=(6, 0))

    def browse_input(self) -> None:
        path = filedialog.askopenfilename(
            title="Select a file",
            filetypes=[
                ("Supported files", "*.txt *.pdf *.docx *.doc"),
                ("Text files", "*.txt"),
                ("PDF files", "*.pdf"),
                ("Word files", "*.docx *.doc"),
                ("All files", "*.*"),
            ],
        )
        if not path:
            return

        input_path = Path(path)
        self.input_var.set(str(input_path))
        self.output_var.set(str(default_output_path(input_path)))
        self.status_var.set("File selected. Enter the word or phrase to redact.")

    def browse_output(self) -> None:
        input_text = self.input_var.get().strip()
        input_path = Path(input_text) if input_text else None

        initial_path = Path(self.output_var.get().strip()) if self.output_var.get().strip() else None
        if initial_path is None and input_path is not None:
            initial_path = default_output_path(input_path)

        defaultextension = ".txt"
        if input_path is not None and input_path.suffix.lower() in SUPPORTED_SUFFIXES:
            defaultextension = default_output_path(input_path).suffix

        path = filedialog.asksaveasfilename(
            title="Choose where to save the redacted file",
            defaultextension=defaultextension,
            initialfile=initial_path.name if initial_path else f"redacted{defaultextension}",
            filetypes=[
                ("Supported outputs", "*.txt *.pdf *.docx"),
                ("Text files", "*.txt"),
                ("PDF files", "*.pdf"),
                ("Word files", "*.docx"),
                ("All files", "*.*"),
            ],
        )
        if path:
            self.output_var.set(path)

    def redact_file(self) -> None:
        input_text = self.input_var.get().strip()
        target = self.target_var.get()
        output_text = self.output_var.get().strip()

        if not input_text:
            messagebox.showerror("Missing file", "Please choose a file first.")
            return

        if not target:
            messagebox.showerror("Missing word", "Please enter a word or phrase to redact.")
            return

        input_path = Path(input_text)
        suffix = input_path.suffix.lower()

        if suffix not in SUPPORTED_SUFFIXES or not input_path.is_file():
            messagebox.showerror(
                "Invalid file",
                "Please choose an existing .txt, .pdf, .docx, or .doc file.",
            )
            return

        output_path = Path(output_text) if output_text else default_output_path(input_path)
        if suffix == ".doc" and output_path.suffix.lower() != ".docx":
            output_path = output_path.with_suffix(".docx")
            self.output_var.set(str(output_path))

        try:
            actual_output_path, count = redact_supported_file(
                input_path=input_path,
                output_path=output_path,
                target=target,
                whole_word=self.whole_word_var.get(),
                ignore_case=not self.case_sensitive_var.get(),
            )
        except UnicodeDecodeError:
            messagebox.showerror("Encoding error", "This text file could not be read as UTF-8.")
            return
        except RuntimeError as exc:
            messagebox.showerror("Dependency or format error", str(exc))
            return
        except OSError as exc:
            messagebox.showerror("File error", str(exc))
            return

        self.output_var.set(str(actual_output_path))
        self.status_var.set(
            f"Done. Redacted {count} occurrence(s) and saved: {actual_output_path}"
        )
        messagebox.showinfo(
            "Redaction complete",
            f"Redacted {count} occurrence(s).\nSaved to:\n{actual_output_path}",
        )


def main() -> None:
    enable_high_dpi()
    root = tk.Tk()
    RedactorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
