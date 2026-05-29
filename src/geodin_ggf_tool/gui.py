from __future__ import annotations

from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .backend import (
    ProcessSettings,
    update_ggf_dates,
    get_excel_headers,
    auto_detect_column,
    DEFAULT_DATE_COLUMN,
    DEFAULT_FILTER_COLUMN,
    DEFAULT_FILTER_MIN,
    DEFAULT_FILTER_MAX,
    DEFAULT_DUMMY_DATE,
    DEFAULT_MAX_SLOTS,
)
from .report import default_report_path, build_user_summary


class GeoDINDateToolApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("GeoDIN GGF Date Tool")
        self.root.geometry("760x620")

        self.ggf_path = tk.StringVar()
        self.xlsx_path = tk.StringVar()
        self.output_path = tk.StringVar()

        self.date_column = tk.StringVar(value=DEFAULT_DATE_COLUMN)
        self.filter_column = tk.StringVar(value=DEFAULT_FILTER_COLUMN)
        self.filter_min = tk.StringVar(value=str(DEFAULT_FILTER_MIN))
        self.filter_max = tk.StringVar(value=str(DEFAULT_FILTER_MAX))
        self.dummy_date = tk.StringVar(value=DEFAULT_DUMMY_DATE)
        self.remove_duplicates = tk.BooleanVar(value=False)

        self.headers: list[str] = []

        self._build_ui()

    def _build_ui(self):
        padding = {"padx": 8, "pady": 6}

        main = ttk.Frame(self.root)
        main.pack(fill="both", expand=True, padx=12, pady=12)

        title = ttk.Label(
            main,
            text="GeoDIN GGF Date Tool",
            font=("Segoe UI", 16, "bold"),
        )
        title.grid(row=0, column=0, columnspan=3, sticky="w", **padding)

        ttk.Label(main, text="GGF template file:").grid(row=1, column=0, sticky="w", **padding)
        ttk.Entry(main, textvariable=self.ggf_path, width=72).grid(row=1, column=1, sticky="we", **padding)
        ttk.Button(main, text="Browse", command=self.select_ggf).grid(row=1, column=2, **padding)

        ttk.Label(main, text="Excel file:").grid(row=2, column=0, sticky="w", **padding)
        ttk.Entry(main, textvariable=self.xlsx_path, width=72).grid(row=2, column=1, sticky="we", **padding)
        ttk.Button(main, text="Browse", command=self.select_xlsx).grid(row=2, column=2, **padding)

        ttk.Label(main, text="Output GGF file:").grid(row=3, column=0, sticky="w", **padding)
        ttk.Entry(main, textvariable=self.output_path, width=72).grid(row=3, column=1, sticky="we", **padding)
        ttk.Button(main, text="Save As", command=self.select_output).grid(row=3, column=2, **padding)

        separator = ttk.Separator(main)
        separator.grid(row=4, column=0, columnspan=3, sticky="we", pady=12)

        ttk.Label(main, text="Date column:").grid(row=5, column=0, sticky="w", **padding)
        self.date_combo = ttk.Combobox(main, textvariable=self.date_column, values=self.headers, width=40)
        self.date_combo.grid(row=5, column=1, sticky="w", **padding)

        ttk.Label(main, text="Filter column:").grid(row=6, column=0, sticky="w", **padding)
        self.filter_combo = ttk.Combobox(main, textvariable=self.filter_column, values=self.headers, width=40)
        self.filter_combo.grid(row=6, column=1, sticky="w", **padding)

        ttk.Label(main, text="Filter min:").grid(row=7, column=0, sticky="w", **padding)
        ttk.Entry(main, textvariable=self.filter_min, width=16).grid(row=7, column=1, sticky="w", **padding)

        ttk.Label(main, text="Filter max:").grid(row=8, column=0, sticky="w", **padding)
        ttk.Entry(main, textvariable=self.filter_max, width=16).grid(row=8, column=1, sticky="w", **padding)

        ttk.Label(main, text="Dummy date:").grid(row=9, column=0, sticky="w", **padding)
        ttk.Entry(main, textvariable=self.dummy_date, width=16).grid(row=9, column=1, sticky="w", **padding)

        ttk.Checkbutton(
            main,
            text="Remove duplicate dates before writing",
            variable=self.remove_duplicates,
        ).grid(row=10, column=1, sticky="w", **padding)

        ttk.Button(
            main,
            text="Update GGF Dates",
            command=self.run_update,
        ).grid(row=11, column=1, sticky="w", padx=8, pady=14)

        ttk.Label(main, text="Log:").grid(row=12, column=0, sticky="nw", **padding)

        self.log_box = tk.Text(main, height=15, width=90)
        self.log_box.grid(row=12, column=1, columnspan=2, sticky="nsew", **padding)

        main.columnconfigure(1, weight=1)
        main.rowconfigure(12, weight=1)

    def log(self, text: str):
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")

    def clear_log(self):
        self.log_box.delete("1.0", "end")

    def select_ggf(self):
        path = filedialog.askopenfilename(
            title="Select GeoDIN GGF template",
            filetypes=[("GeoDIN GGF files", "*.GGF *.ggf"), ("All files", "*.*")]
        )
        if path:
            self.ggf_path.set(path)

            if not self.output_path.get():
                p = Path(path)
                self.output_path.set(str(p.with_name(f"{p.stem}_updated.GGF")))

    def select_xlsx(self):
        path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if path:
            self.xlsx_path.set(path)
            self.load_headers(Path(path))

    def select_output(self):
        path = filedialog.asksaveasfilename(
            title="Save updated GGF as",
            defaultextension=".GGF",
            filetypes=[("GeoDIN GGF files", "*.GGF *.ggf"), ("All files", "*.*")]
        )
        if path:
            self.output_path.set(path)

    def load_headers(self, xlsx_path: Path):
        try:
            self.headers = get_excel_headers(xlsx_path)
            self.date_combo["values"] = self.headers
            self.filter_combo["values"] = self.headers

            detected_date = auto_detect_column(self.headers, DEFAULT_DATE_COLUMN)
            detected_filter = auto_detect_column(self.headers, DEFAULT_FILTER_COLUMN)

            if detected_date:
                self.date_column.set(detected_date)

            if detected_filter:
                self.filter_column.set(detected_filter)

            self.log(f"Loaded Excel columns: {', '.join(self.headers)}")

        except Exception as exc:
            messagebox.showerror("Excel error", str(exc))

    def validate_inputs(self) -> bool:
        if not self.ggf_path.get():
            messagebox.showerror("Missing input", "Please select a GGF template file.")
            return False

        if not self.xlsx_path.get():
            messagebox.showerror("Missing input", "Please select an Excel file.")
            return False

        if not self.output_path.get():
            messagebox.showerror("Missing output", "Please select an output GGF file.")
            return False

        try:
            float(self.filter_min.get())
            float(self.filter_max.get())
        except ValueError:
            messagebox.showerror("Invalid filter range", "Filter min and max must be numbers.")
            return False

        dummy = self.dummy_date.get().strip()
        if len(dummy) != 8 or not dummy.isdigit():
            messagebox.showerror("Invalid dummy date", "Dummy date must be 8 digits, e.g. 19000101.")
            return False

        return True

    def run_update(self):
        if not self.validate_inputs():
            return

        self.clear_log()

        output_path = Path(self.output_path.get())
        report_path = default_report_path(output_path)

        settings = ProcessSettings(
            ggf_input=Path(self.ggf_path.get()),
            xlsx_input=Path(self.xlsx_path.get()),
            ggf_output=output_path,
            report_output=report_path,
            date_column=self.date_column.get(),
            filter_column=self.filter_column.get(),
            filter_min=float(self.filter_min.get()),
            filter_max=float(self.filter_max.get()),
            dummy_date=self.dummy_date.get().strip(),
            max_slots=DEFAULT_MAX_SLOTS,
            remove_duplicate_dates=self.remove_duplicates.get(),
        )

        try:
            result = update_ggf_dates(settings)

            self.log("\n".join(result.log_lines))
            self.log("")
            self.log(f"Report saved: {report_path}")

            if result.duplicate_dates and not self.remove_duplicates.get():
                messagebox.showwarning(
                    "Duplicate dates detected",
                    "Duplicate dates were detected.\n\n"
                    "The file was still created.\n"
                    "Tick 'Remove duplicate dates before writing' if you want to remove them."
                )

            messagebox.showinfo(
                "Finished",
                build_user_summary(result) + f"\nReport saved:\n{report_path}"
            )

        except Exception as exc:
            self.log(f"ERROR: {exc}")
            messagebox.showerror("Processing error", str(exc))


def run_app():
    root = tk.Tk()
    app = GeoDINDateToolApp(root)
    root.mainloop()
