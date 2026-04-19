#!/usr/bin/env python3
from __future__ import annotations

from pathlib import Path
import re
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

BOX = "\u2588"


def default_output_path(input_path: Path) -> Path:
    return input_path.with_name(f"{input_path.stem}_redacted.txt")


def redact_text(text: str, target: str, whole_word: bool, ignore_case: bool) -> tuple[str, int]:
    pattern = re.escape(target)
    if whole_word:
        pattern = rf"\b{pattern}\b"

    flags = re.IGNORECASE if ignore_case else 0
    regex = re.compile(pattern, flags)

    def replacer(match: re.Match[str]) -> str:
        return BOX * len(match.group(0))

    return regex.subn(replacer, text)


class RedactorApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("TXT File Redactor")
        self.root.resizable(False, False)

        self.input_var = tk.StringVar()
        self.output_var = tk.StringVar()
        self.target_var = tk.StringVar()
        self.whole_word_var = tk.BooleanVar(value=True)
        self.case_sensitive_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="Choose a text file to begin.")

        self.build_ui()

    def build_ui(self) -> None:
        frame = ttk.Frame(self.root, padding=14)
        frame.grid(sticky="nsew")
        frame.columnconfigure(1, weight=1)

        ttk.Label(frame, text="Text file").grid(row=0, column=0, padx=10, pady=6, sticky="w")
        ttk.Entry(frame, textvariable=self.input_var, width=48).grid(row=0, column=1, padx=10, pady=6, sticky="ew")
        ttk.Button(frame, text="Browse", command=self.browse_input).grid(row=0, column=2, padx=10, pady=6)

        ttk.Label(frame, text="Word or phrase").grid(row=1, column=0, padx=10, pady=6, sticky="w")
        ttk.Entry(frame, textvariable=self.target_var, width=48).grid(row=1, column=1, columnspan=2, padx=10, pady=6, sticky="ew")

        ttk.Label(frame, text="Save as").grid(row=2, column=0, padx=10, pady=6, sticky="w")
        ttk.Entry(frame, textvariable=self.output_var, width=48).grid(row=2, column=1, padx=10, pady=6, sticky="ew")
        ttk.Button(frame, text="Browse", command=self.browse_output).grid(row=2, column=2, padx=10, pady=6)

        ttk.Checkbutton(frame, text="Whole word only", variable=self.whole_word_var).grid(row=3, column=0, columnspan=2, padx=10, pady=6, sticky="w")
        ttk.Checkbutton(frame, text="Case sensitive", variable=self.case_sensitive_var).grid(row=4, column=0, columnspan=2, padx=10, pady=6, sticky="w")

        ttk.Button(frame, text="Redact File", command=self.redact_file).grid(row=5, column=0, columnspan=3, padx=10, pady=(10, 6), sticky="ew")
        ttk.Label(frame, textvariable=self.status_var, wraplength=420, justify="left").grid(row=6, column=0, columnspan=3, padx=10, pady=(6, 0), sticky="w")

    def browse_input(self) -> None:
        path = filedialog.askopenfilename(
            title="Select a text file",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
        )
        if not path:
            return

        input_path = Path(path)
        self.input_var.set(str(input_path))
        if not self.output_var.get().strip():
            self.output_var.set(str(default_output_path(input_path)))
        self.status_var.set("File selected. Enter the word or phrase to redact.")

    def browse_output(self) -> None:
        initial = self.output_var.get().strip()
        path = filedialog.asksaveasfilename(
            title="Choose where to save the redacted file",
            defaultextension=".txt",
            initialfile=Path(initial).name if initial else "redacted.txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
        )
        if path:
            self.output_var.set(path)

    def redact_file(self) -> None:
        input_text = self.input_var.get().strip()
        target = self.target_var.get()
        output_text = self.output_var.get().strip()

        if not input_text:
            messagebox.showerror("Missing file", "Please choose a .txt file first.")
            return

        if not target:
            messagebox.showerror("Missing word", "Please enter a word or phrase to redact.")
            return

        input_path = Path(input_text)
        if input_path.suffix.lower() != ".txt" or not input_path.is_file():
            messagebox.showerror("Invalid file", "Please choose an existing .txt file.")
            return

        output_path = Path(output_text) if output_text else default_output_path(input_path)
        self.output_var.set(str(output_path))

        try:
            text = input_path.read_text(encoding="utf-8")
            redacted_text, count = redact_text(
                text=text,
                target=target,
                whole_word=self.whole_word_var.get(),
                ignore_case=not self.case_sensitive_var.get(),
            )
            output_path.write_text(redacted_text, encoding="utf-8")
        except UnicodeDecodeError:
            messagebox.showerror("Encoding error", "This file could not be read as UTF-8 text.")
            return
        except OSError as exc:
            messagebox.showerror("File error", str(exc))
            return

        self.status_var.set(f"Done. Redacted {count} occurrence(s) and saved: {output_path}")
        messagebox.showinfo("Redaction complete", f"Redacted {count} occurrence(s).\nSaved to:\n{output_path}")


def main() -> None:
    root = tk.Tk()
    RedactorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
