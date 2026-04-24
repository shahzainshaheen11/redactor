#!/usr/bin/env python3
"""CLI redactor for TXT, DOCX, DOC, and PDF files."""

from __future__ import annotations

import argparse
from pathlib import Path
import re
import tempfile

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

    raise RuntimeError("Unsupported file type. Use .txt, .pdf, .docx, or .doc.")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Replace a selected word or phrase in TXT, PDF, DOCX, or DOC files "
            "with black box characters of the same length."
        )
    )
    parser.add_argument("input_file", type=Path, help="Path to the input file")
    parser.add_argument("target", help="Word or phrase to redact")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        help="Optional output file path. Defaults to '<input_stem>_redacted.<ext>'",
    )
    parser.add_argument(
        "--whole-word",
        action="store_true",
        help="Only redact whole-word matches",
    )
    parser.add_argument(
        "--case-sensitive",
        action="store_true",
        help="Match case exactly instead of ignoring case",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    input_path = args.input_file

    if input_path.suffix.lower() not in SUPPORTED_SUFFIXES:
        raise SystemExit("Input file must be .txt, .pdf, .docx, or .doc.")

    if not input_path.is_file():
        raise SystemExit(f"File not found: {input_path}")

    output_path = args.output or default_output_path(input_path)

    try:
        actual_output_path, count = redact_supported_file(
            input_path=input_path,
            output_path=output_path,
            target=args.target,
            whole_word=args.whole_word,
            ignore_case=not args.case_sensitive,
        )
    except RuntimeError as exc:
        raise SystemExit(str(exc)) from exc

    print(f"Input: {input_path}")
    print(f"Output: {actual_output_path}")
    print(f"Redacted occurrences: {count}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
