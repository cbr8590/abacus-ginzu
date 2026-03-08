#!/usr/bin/env python3
"""
Ginzu Autocomplete - Pop-up desktop app.
Drag-and-drop or browse for documents, pick template, run, open output.
"""

import os
import subprocess
import sys
import threading
from pathlib import Path
from typing import Optional

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

# Ensure we can import from this directory
sys.path.insert(0, str(Path(__file__).resolve().parent))

import ginzu_debug
import tkinter as tk
from tkinter import filedialog, messagebox

# Default template path
DEFAULT_TEMPLATE = os.path.expanduser("~/Downloads/higrowth.xls")

# Modern dark theme
BG = "#0d1117"
SURFACE = "#161b22"
SURFACE_HOVER = "#21262d"
BORDER = "#30363d"
FG = "#e6edf3"
FG_MUTED = "#8b949e"
ACCENT = "#58a6ff"
ACCENT_HOVER = "#79b8ff"
SUCCESS = "#3fb950"
ENTRY_BG = "#21262d"
BTN_SECONDARY = "#21262d"
FONT_FAMILY = ("SF Pro Display", "Helvetica Neue", "Helvetica", "Arial")
FONT_SIZE = 11
FONT_MONO = ("SF Mono", "Menlo", "Monaco", "Consolas", "monospace")


def run_autocomplete_sync(
    document_path=None,
    company_name=None,
    template_path=None,
    output_path=None,
    instructions_path=None,
):
    """Import and call run_autocomplete from autocomplete_ginzu."""
    from autocomplete_ginzu import run_autocomplete

    return run_autocomplete(
        document_path=document_path,
        company_name=company_name or None,
        template_path=template_path or DEFAULT_TEMPLATE,
        output_path=output_path,
        instructions_path=instructions_path,
    )


class GinzuApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Ginzu Model Autocomplete")
        self.root.geometry("520x680")
        self.root.resizable(True, True)
        self.root.minsize(440, 560)

        self.root.configure(bg=BG)
        self.root.tk_setPalette(background=BG, foreground=FG)

        self.stay_on_top = tk.BooleanVar(value=False)
        self.root.attributes("-topmost", False)

        self.output_path = None
        self._build_ui()
        self._setup_debug_handler()
        self._log("Ready. Select a document or enter a company name, then click Run.")
        self.root.update()

    def _section(self, parent, title: str) -> tk.Frame:
        """Create a card-like section with title."""
        frame = tk.Frame(parent, bg=SURFACE, highlightbackground=BORDER, highlightthickness=1)
        frame.pack(fill=tk.X, pady=10)
        header = tk.Frame(frame, bg=SURFACE)
        header.pack(fill=tk.X, padx=14, pady=8)
        tk.Label(
            header, text=title, font=(FONT_FAMILY[0], 10, "bold"),
            bg=SURFACE, fg=FG_MUTED,
        ).pack(anchor=tk.W)
        inner = tk.Frame(frame, bg=SURFACE)
        inner.pack(fill=tk.X, padx=14, pady=12)
        return inner

    def _row(self, parent, label: str, entry_var: tk.StringVar, entry_widget, browse_cmd=None) -> tk.Frame:
        """Create a labeled row with optional Browse button."""
        row = tk.Frame(parent, bg=SURFACE)
        row.pack(fill=tk.X, pady=6)
        tk.Label(row, text=label, width=10, anchor=tk.W, font=(FONT_FAMILY[0], FONT_SIZE),
                 bg=SURFACE, fg=FG_MUTED).pack(side=tk.LEFT, padx=8)
        entry_widget.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=8)
        if browse_cmd:
            btn = tk.Button(row, text="Browse", command=browse_cmd, font=(FONT_FAMILY[0], 10),
                            bg=BTN_SECONDARY, fg=FG, activebackground=SURFACE_HOVER, activeforeground=FG,
                            relief=tk.FLAT, padx=12, pady=4, cursor="hand2")
            btn.pack(side=tk.LEFT)
        return row

    def _build_ui(self):
        main = tk.Frame(self.root, bg=BG)
        main.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        # Header
        header = tk.Frame(main, bg=BG)
        header.pack(fill=tk.X, pady=16)
        tk.Label(header, text="Ginzu", font=(FONT_FAMILY[0], 22, "bold"),
                 bg=BG, fg=FG).pack(anchor=tk.W)
        tk.Label(header, text="Fill Damodaran valuation model from documents or company data",
                 font=(FONT_FAMILY[0], 11), bg=BG, fg=FG_MUTED).pack(anchor=tk.W)

        # Mode section
        mode_inner = self._section(main, "MODE")
        self.mode = tk.StringVar(value="document")
        for val, txt in [
            ("document", "Document — Extract from PDF/Word diligence"),
            ("company", "Company — Infer from name or ticker"),
        ]:
            rb = tk.Radiobutton(
                mode_inner, text=txt, variable=self.mode, value=val, command=self._on_mode_change,
                font=(FONT_FAMILY[0], FONT_SIZE), bg=SURFACE, fg=FG,
                selectcolor=SURFACE, activebackground=SURFACE, activeforeground=FG,
            )
            rb.pack(anchor=tk.W, pady=4)

        # Inputs section
        inputs_inner = self._section(main, "INPUTS")
        self.doc_var = tk.StringVar()
        self.doc_entry = tk.Entry(inputs_inner, textvariable=self.doc_var, width=32,
                                  font=(FONT_MONO[0], 10), bg=ENTRY_BG, fg=FG, insertbackground=FG,
                                  relief=tk.FLAT, highlightthickness=1, highlightbackground=BORDER)
        self._row(inputs_inner, "Document", self.doc_var, self.doc_entry, self._browse_document)

        self.company_var = tk.StringVar()
        self.company_entry = tk.Entry(inputs_inner, textvariable=self.company_var, width=32,
                                      font=(FONT_MONO[0], 10), bg=ENTRY_BG, fg=FG, insertbackground=FG,
                                      relief=tk.FLAT, highlightthickness=1, highlightbackground=BORDER)
        self._row(inputs_inner, "Company", self.company_var, self.company_entry)
        tk.Label(inputs_inner, text="Optional for document mode", font=(FONT_FAMILY[0], 9),
                 bg=SURFACE, fg=FG_MUTED).pack(anchor=tk.W, padx=82, pady=4)

        self.tpl_var = tk.StringVar(value=DEFAULT_TEMPLATE)
        self.tpl_entry = tk.Entry(inputs_inner, textvariable=self.tpl_var, width=32,
                                  font=(FONT_MONO[0], 10), bg=ENTRY_BG, fg=FG, insertbackground=FG,
                                  relief=tk.FLAT, highlightthickness=1, highlightbackground=BORDER)
        self._row(inputs_inner, "Template", self.tpl_var, self.tpl_entry, self._browse_template)

        # Actions
        btn_frame = tk.Frame(main, bg=BG)
        btn_frame.pack(fill=tk.X, pady=12)
        self.run_btn = tk.Button(btn_frame, text="Run", command=self._run, font=(FONT_FAMILY[0], 11, "bold"),
                                bg=ACCENT, fg="#ffffff", activebackground=ACCENT_HOVER, activeforeground="#ffffff",
                                relief=tk.FLAT, padx=24, pady=10, cursor="hand2")
        self.run_btn.pack(side=tk.LEFT, padx=8)
        self.open_btn = tk.Button(btn_frame, text="Open output", command=self._open_output,
                                 font=(FONT_FAMILY[0], 10), state=tk.DISABLED,
                                 bg=BTN_SECONDARY, fg=FG, activebackground=SURFACE_HOVER, activeforeground=FG,
                                 relief=tk.FLAT, padx=16, pady=8, cursor="hand2")
        self.open_btn.pack(side=tk.LEFT)

        # Status
        status_inner = self._section(main, "STATUS")
        self.status_text = tk.Text(status_inner, height=2, wrap=tk.WORD, state=tk.DISABLED,
                                   font=(FONT_MONO[0], 10), bg=SURFACE, fg=FG, insertbackground=FG)
        self.status_text.pack(fill=tk.X)

        # Debug
        debug_inner = self._section(main, "LOG")
        debug_row = tk.Frame(debug_inner, bg=SURFACE)
        debug_row.pack(fill=tk.BOTH, expand=True)
        self.debug_text = tk.Text(debug_row, height=10, wrap=tk.WORD, font=(FONT_MONO[0], 9),
                                  bg=SURFACE, fg=FG_MUTED, insertbackground=FG)
        self.debug_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll = tk.Scrollbar(debug_row, command=self.debug_text.yview, bg=SURFACE)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.debug_text.config(yscrollcommand=scroll.set)
        self.debug_text.insert(tk.END, "Ready. Click Run to see activity.\n")
        tk.Button(debug_inner, text="Clear", command=self._clear_debug, font=(FONT_FAMILY[0], 9),
                  bg=BTN_SECONDARY, fg=FG_MUTED, relief=tk.FLAT, padx=10, pady=4,
                  cursor="hand2").pack(anchor=tk.E, pady=6)

        # Footer
        footer = tk.Frame(main, bg=BG)
        footer.pack(fill=tk.X, pady=8)
        tk.Checkbutton(
            footer, text="Keep on top", variable=self.stay_on_top, command=self._toggle_topmost,
            font=(FONT_FAMILY[0], 10), bg=BG, fg=FG_MUTED, selectcolor=SURFACE,
            activebackground=BG, activeforeground=FG,
        ).pack(side=tk.LEFT)

        self._on_mode_change()

    def _on_mode_change(self):
        if self.mode.get() == "document":
            self.doc_entry.config(state=tk.NORMAL)
            self.company_entry.config(state=tk.NORMAL)
        else:
            self.doc_var.set("")
            self.doc_entry.config(state=tk.DISABLED)
            self.company_entry.config(state=tk.NORMAL)

    def _toggle_topmost(self):
        self.root.attributes("-topmost", self.stay_on_top.get())

    def _browse_document(self):
        path = filedialog.askopenfilename(
            title="Select company diligence document",
            filetypes=[
                ("PDF files", "*.pdf"),
                ("Word files", "*.docx"),
                ("Text files", "*.txt"),
            ],
        )
        if path:
            self.doc_var.set(path)

    def _browse_template(self):
        path = filedialog.askopenfilename(
            title="Select Ginzu Excel template",
            filetypes=[
                ("Excel files", "*.xls"),
                ("Excel files", "*.xlsx"),
            ],
        )
        if path:
            self.tpl_var.set(path)

    def _log(self, msg: str):
        self.status_text.config(state=tk.NORMAL)
        self.status_text.insert(tk.END, msg + "\n")
        self.status_text.see(tk.END)
        self.status_text.config(state=tk.DISABLED)
        self.root.update_idletasks()

    def _clear_status(self):
        self.status_text.config(state=tk.NORMAL)
        self.status_text.delete(1.0, tk.END)
        self.status_text.config(state=tk.DISABLED)

    def _clear_debug(self):
        self.debug_text.config(state=tk.NORMAL)
        self.debug_text.delete(1.0, tk.END)
        self.debug_text.config(state=tk.NORMAL)

    def _setup_debug_handler(self):
        root, debug_text = self.root, self.debug_text

        def handler(msg: str, level: str = "info"):
            def append():
                debug_text.config(state=tk.NORMAL)
                prefix = f"[{level.upper():5}] " if level != "info" else ""
                debug_text.insert(tk.END, f"{prefix}{msg}\n")
                debug_text.see(tk.END)
                debug_text.config(state=tk.NORMAL)
            root.after(0, append)

        ginzu_debug.set_handler(handler)
        ginzu_debug.log("Ready. Click Run to see activity.", "info")

    def _run(self):
        mode = self.mode.get()
        doc_path = self.doc_var.get().strip() or None
        company = self.company_var.get().strip() or None
        tpl_path = self.tpl_var.get().strip() or DEFAULT_TEMPLATE

        if mode == "document":
            if not doc_path:
                messagebox.showerror("Error", "Please select a document.")
                return
        else:
            if not company:
                messagebox.showerror("Error", "Please enter a company name or ticker.")
                return

        if not Path(tpl_path).exists():
            messagebox.showerror("Error", f"Template not found:\n{tpl_path}")
            return

        self._clear_status()
        self.run_btn.config(state=tk.DISABLED)
        self.open_btn.config(state=tk.DISABLED)
        self.output_path = None

        def work():
            try:
                ginzu_debug.log("=== Run started ===", "info")
                ginzu_debug.log(f"Mode: {mode} | Doc: {doc_path or 'N/A'} | Company: {company or 'N/A'}", "info")
                self._log("Running…")
                success, out_path, err = run_autocomplete_sync(
                    document_path=doc_path if mode == "document" else None,
                    company_name=company if mode == "company" else (company or None),
                    template_path=tpl_path,
                )
                self.root.after(0, lambda: self._on_done(success, out_path, err))
            except Exception as e:
                self.root.after(0, lambda: self._on_done(False, None, str(e)))

        threading.Thread(target=work, daemon=True).start()

    def _on_done(self, success: bool, out_path: Optional[str], error: Optional[str]):
        self.run_btn.config(state=tk.NORMAL)
        if success and out_path:
            self.output_path = out_path
            self.open_btn.config(state=tk.NORMAL)
            self._log(f"Done. Output: {out_path}")
        else:
            self._log(f"Error: {error or 'Unknown error'}")
            messagebox.showerror("Error", error or "Unknown error")

    def _open_output(self):
        if self.output_path and Path(self.output_path).exists():
            if sys.platform == "darwin":
                subprocess.run(["open", self.output_path], check=False)
            elif sys.platform == "win32":
                os.startfile(self.output_path)
            else:
                subprocess.run(["xdg-open", self.output_path], check=False)
        else:
            messagebox.showwarning("Warning", "Output file not found.")

    def run(self):
        self.root.mainloop()


def main():
    app = GinzuApp()
    app.run()


if __name__ == "__main__":
    main()
