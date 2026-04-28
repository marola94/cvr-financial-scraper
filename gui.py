"""
CVR Lookup Tool — GUI entry point.
Double-click cvr_lookup.exe to launch.
"""
import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk

import main as cvr_main


# ── Stdout redirector ──────────────────────────────────────────────────────────

class _TextSink:
    def __init__(self, widget: scrolledtext.ScrolledText):
        self._w = widget

    def write(self, text: str):
        self._w.after(0, self._append, text)

    def _append(self, text: str):
        self._w.configure(state="normal")
        self._w.insert(tk.END, text)
        self._w.see(tk.END)
        self._w.configure(state="disabled")

    def flush(self):
        pass


# ── Main application ───────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("CVR Lookup Tool")
        self.resizable(True, True)
        self.minsize(640, 480)
        self._output_path: str | None = None
        self._build_ui()
        self.geometry("760x560")

    # ── UI construction ────────────────────────────────────────────────────────

    def _build_ui(self):
        pad = {"padx": 12, "pady": 6}

        # ── Header ────────────────────────────────────────────────────────────
        hdr = tk.Frame(self, bg="#1a3a5c")
        hdr.pack(fill="x")
        tk.Label(
            hdr, text="CVR Lookup Tool",
            bg="#1a3a5c", fg="white",
            font=("Segoe UI", 16, "bold"),
            pady=12,
        ).pack(side="left", padx=16)

        # ── File selection ────────────────────────────────────────────────────
        frame_files = tk.LabelFrame(self, text="Filer", font=("Segoe UI", 10))
        frame_files.pack(fill="x", **pad)

        # Input file
        tk.Label(frame_files, text="Input fil:", width=10, anchor="w").grid(
            row=0, column=0, padx=8, pady=4, sticky="w"
        )
        self._input_var = tk.StringVar()
        tk.Entry(frame_files, textvariable=self._input_var, width=55).grid(
            row=0, column=1, padx=4, pady=4, sticky="ew"
        )
        tk.Button(
            frame_files, text="Vælg…", command=self._browse_input, width=8
        ).grid(row=0, column=2, padx=8, pady=4)

        # Output file
        tk.Label(frame_files, text="Output fil:", width=10, anchor="w").grid(
            row=1, column=0, padx=8, pady=4, sticky="w"
        )
        self._output_var = tk.StringVar(value="output.xlsx")
        tk.Entry(frame_files, textvariable=self._output_var, width=55).grid(
            row=1, column=1, padx=4, pady=4, sticky="ew"
        )
        tk.Button(
            frame_files, text="Vælg…", command=self._browse_output, width=8
        ).grid(row=1, column=2, padx=8, pady=4)

        frame_files.columnconfigure(1, weight=1)

        # ── Run button ────────────────────────────────────────────────────────
        btn_frame = tk.Frame(self)
        btn_frame.pack(fill="x", padx=12, pady=4)

        self._run_btn = tk.Button(
            btn_frame,
            text="▶  Kør opslag",
            command=self._start_run,
            bg="#1a7a3c", fg="white",
            font=("Segoe UI", 11, "bold"),
            padx=16, pady=6,
            relief="flat",
            cursor="hand2",
        )
        self._run_btn.pack(side="left")

        self._open_btn = tk.Button(
            btn_frame,
            text="📂  Åbn output",
            command=self._open_output,
            bg="#1a3a5c", fg="white",
            font=("Segoe UI", 11),
            padx=16, pady=6,
            relief="flat",
            cursor="hand2",
        )
        # Hidden until output is ready

        # ── Progress bar ──────────────────────────────────────────────────────
        self._progress = ttk.Progressbar(self, mode="indeterminate")
        self._progress.pack(fill="x", padx=12, pady=(0, 4))

        # ── Log area ──────────────────────────────────────────────────────────
        log_frame = tk.LabelFrame(self, text="Log", font=("Segoe UI", 10))
        log_frame.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        self._log = scrolledtext.ScrolledText(
            log_frame,
            font=("Consolas", 9),
            state="disabled",
            bg="#f7f7f7",
            wrap="word",
        )
        self._log.pack(fill="both", expand=True, padx=4, pady=4)

    # ── File dialogs ──────────────────────────────────────────────────────────

    def _browse_input(self):
        path = filedialog.askopenfilename(
            title="Vælg input fil",
            filetypes=[("Excel / CSV", "*.xlsx *.xls *.csv"), ("Alle filer", "*.*")],
        )
        if path:
            self._input_var.set(path)
            # Auto-suggest output in same folder
            folder = os.path.dirname(path)
            self._output_var.set(os.path.join(folder, "output.xlsx"))

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title="Gem output som",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=self._output_var.get(),
        )
        if path:
            self._output_var.set(path)

    # ── Run logic ─────────────────────────────────────────────────────────────

    def _start_run(self):
        input_file = self._input_var.get().strip()
        output_file = self._output_var.get().strip()

        if not input_file:
            self._log_line("⚠  Vælg en input fil først.\n")
            return
        if not os.path.exists(input_file):
            self._log_line(f"⚠  Filen findes ikke: {input_file}\n")
            return

        # Make output path absolute (relative to input folder)
        if not os.path.isabs(output_file):
            output_file = os.path.join(os.path.dirname(input_file), output_file)
            self._output_var.set(output_file)

        self._output_path = output_file
        self._open_btn.pack_forget()
        self._run_btn.configure(state="disabled")
        self._progress.start(12)
        self._clear_log()

        thread = threading.Thread(
            target=self._run_task, args=(input_file, output_file), daemon=True
        )
        thread.start()

    def _run_task(self, input_file: str, output_file: str):
        old_stdout = sys.stdout
        old_stderr = sys.stderr
        sink = _TextSink(self._log)
        sys.stdout = sink
        sys.stderr = sink

        try:
            cvr_main.main(input_file, output_file)
        except SystemExit:
            pass
        except Exception as exc:
            print(f"\n[FEJL] {exc}")
        finally:
            sys.stdout = old_stdout
            sys.stderr = old_stderr
            self.after(0, self._run_done)

    def _run_done(self):
        self._progress.stop()
        self._run_btn.configure(state="normal")
        if self._output_path and os.path.exists(self._output_path):
            self._open_btn.pack(side="left", padx=(12, 0))

    # ── Helpers ───────────────────────────────────────────────────────────────

    def _open_output(self):
        if self._output_path and os.path.exists(self._output_path):
            os.startfile(self._output_path)

    def _log_line(self, text: str):
        self._log.configure(state="normal")
        self._log.insert(tk.END, text)
        self._log.see(tk.END)
        self._log.configure(state="disabled")

    def _clear_log(self):
        self._log.configure(state="normal")
        self._log.delete("1.0", tk.END)
        self._log.configure(state="disabled")


# ── Entry point ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()
