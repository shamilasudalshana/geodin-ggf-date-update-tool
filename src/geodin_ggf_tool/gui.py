from __future__ import annotations

from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_FILES = None
    TkinterDnD = None
    DND_AVAILABLE = False

try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    Image = None
    ImageTk = None
    PIL_AVAILABLE = False

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
from .translations import TEXT


class GeoDINDateToolApp:
    def __init__(self, root: tk.Tk):
        self.root = root

        self.lang = tk.StringVar(value="en")

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
        self.widgets = {}
        self.logo_image = None

        self._build_ui()
        self.load_logo()
        self.apply_language()

        if DND_AVAILABLE:
            self.enable_drag_and_drop()

    def t(self, key: str) -> str:
        return TEXT[self.lang.get()].get(key, key)

    def _build_ui(self):
        self.root.geometry("820x660")

        padding = {"padx": 8, "pady": 6}

        main = ttk.Frame(self.root)
        main.pack(fill="both", expand=True, padx=12, pady=12)

        self.widgets["title"] = ttk.Label(main, font=("Segoe UI", 16, "bold"))
        self.widgets["title"].grid(row=0, column=0, columnspan=2, sticky="w", **padding)

        self.widgets["logo"] = ttk.Label(main)
        self.widgets["logo"].grid(row=0, column=2, sticky="e", **padding)

        self.widgets["language_label"] = ttk.Label(main)
        self.widgets["language_label"].grid(row=1, column=0, sticky="w", **padding)

        lang_box = ttk.Combobox(
            main,
            textvariable=self.lang,
            values=["en", "de"],
            width=8,
            state="readonly",
        )
        lang_box.grid(row=1, column=1, sticky="w", **padding)
        lang_box.bind("<<ComboboxSelected>>", lambda event: self.apply_language())

        self.widgets["ggf_label"] = ttk.Label(main)
        self.widgets["ggf_label"].grid(row=2, column=0, sticky="w", **padding)
        ttk.Entry(main, textvariable=self.ggf_path, width=72).grid(row=2, column=1, sticky="we", **padding)
        self.widgets["ggf_browse"] = ttk.Button(main, command=self.select_ggf)
        self.widgets["ggf_browse"].grid(row=2, column=2, **padding)

        self.widgets["excel_label"] = ttk.Label(main)
        self.widgets["excel_label"].grid(row=3, column=0, sticky="w", **padding)
        ttk.Entry(main, textvariable=self.xlsx_path, width=72).grid(row=3, column=1, sticky="we", **padding)
        self.widgets["excel_browse"] = ttk.Button(main, command=self.select_xlsx)
        self.widgets["excel_browse"].grid(row=3, column=2, **padding)

        self.widgets["output_label"] = ttk.Label(main)
        self.widgets["output_label"].grid(row=4, column=0, sticky="w", **padding)
        ttk.Entry(main, textvariable=self.output_path, width=72).grid(row=4, column=1, sticky="we", **padding)
        self.widgets["output_browse"] = ttk.Button(main, command=self.select_output)
        self.widgets["output_browse"].grid(row=4, column=2, **padding)

        separator = ttk.Separator(main)
        separator.grid(row=5, column=0, columnspan=3, sticky="we", pady=12)

        self.widgets["date_column_label"] = ttk.Label(main)
        self.widgets["date_column_label"].grid(row=6, column=0, sticky="w", **padding)
        self.date_combo = ttk.Combobox(main, textvariable=self.date_column, values=self.headers, width=40)
        self.date_combo.grid(row=6, column=1, sticky="w", **padding)

        self.widgets["filter_column_label"] = ttk.Label(main)
        self.widgets["filter_column_label"].grid(row=7, column=0, sticky="w", **padding)
        self.filter_combo = ttk.Combobox(main, textvariable=self.filter_column, values=self.headers, width=40)
        self.filter_combo.grid(row=7, column=1, sticky="w", **padding)

        self.widgets["filter_min_label"] = ttk.Label(main)
        self.widgets["filter_min_label"].grid(row=8, column=0, sticky="w", **padding)
        ttk.Entry(main, textvariable=self.filter_min, width=16).grid(row=8, column=1, sticky="w", **padding)

        self.widgets["filter_max_label"] = ttk.Label(main)
        self.widgets["filter_max_label"].grid(row=9, column=0, sticky="w", **padding)
        ttk.Entry(main, textvariable=self.filter_max, width=16).grid(row=9, column=1, sticky="w", **padding)

        self.widgets["dummy_date_label"] = ttk.Label(main)
        self.widgets["dummy_date_label"].grid(row=10, column=0, sticky="w", **padding)
        ttk.Entry(main, textvariable=self.dummy_date, width=16).grid(row=10, column=1, sticky="w", **padding)

        self.widgets["remove_duplicates_check"] = ttk.Checkbutton(
            main,
            variable=self.remove_duplicates,
        )
        self.widgets["remove_duplicates_check"].grid(row=11, column=1, sticky="w", **padding)

        self.widgets["update_button"] = ttk.Button(main, command=self.run_update)
        self.widgets["update_button"].grid(row=12, column=1, sticky="w", padx=8, pady=14)

        self.widgets["drop_hint"] = ttk.Label(main, foreground="gray")
        self.widgets["drop_hint"].grid(row=12, column=1, sticky="e", **padding)

        self.widgets["log_label"] = ttk.Label(main)
        self.widgets["log_label"].grid(row=13, column=0, sticky="nw", **padding)

        self.log_box = tk.Text(main, height=15, width=90)
        self.log_box.grid(row=13, column=1, columnspan=2, sticky="nsew", **padding)

        main.columnconfigure(1, weight=1)
        main.rowconfigure(13, weight=1)

    def load_logo(self):
        if not PIL_AVAILABLE:
            return

        possible_paths = [
            Path.cwd() / "assets" / "logo.png",
            Path(__file__).resolve().parents[2] / "assets" / "logo.png",
            Path(__file__).resolve().parents[3] / "assets" / "logo.png",
        ]

        for logo_path in possible_paths:
            if logo_path.exists():
                try:
                    image = Image.open(logo_path)
                    image.thumbnail((70, 70))
                    self.logo_image = ImageTk.PhotoImage(image)
                    self.widgets["logo"].config(image=self.logo_image)
                    return
                except Exception:
                    return

    def apply_language(self):
        self.root.title(self.t("app_title"))

        self.widgets["title"].config(text=self.t("app_title"))
        self.widgets["language_label"].config(text=self.t("language"))
        self.widgets["ggf_label"].config(text=self.t("ggf_template"))
        self.widgets["excel_label"].config(text=self.t("excel_file"))
        self.widgets["output_label"].config(text=self.t("output_file"))

        self.widgets["ggf_browse"].config(text=self.t("browse"))
        self.widgets["excel_browse"].config(text=self.t("browse"))
        self.widgets["output_browse"].config(text=self.t("save_as"))

        self.widgets["date_column_label"].config(text=self.t("date_column"))
        self.widgets["filter_column_label"].config(text=self.t("filter_column"))
        self.widgets["filter_min_label"].config(text=self.t("filter_min"))
        self.widgets["filter_max_label"].config(text=self.t("filter_max"))
        self.widgets["dummy_date_label"].config(text=self.t("dummy_date"))

        self.widgets["remove_duplicates_check"].config(text=self.t("remove_duplicates"))
        self.widgets["update_button"].config(text=self.t("update_button"))
        self.widgets["log_label"].config(text=self.t("log"))
        self.widgets["drop_hint"].config(text=self.t("drop_hint"))

    def enable_drag_and_drop(self):
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind("<<Drop>>", self.handle_drop)

    def handle_drop(self, event):
        raw_files = self.root.tk.splitlist(event.data)

        for raw_file in raw_files:
            path = Path(raw_file)

            if path.suffix.lower() == ".ggf":
                self.ggf_path.set(str(path))

                if not self.output_path.get():
                    self.output_path.set(str(path.with_name(f"{path.stem}_updated.GGF")))

                self.log(f"GGF selected by drag-and-drop: {path}")

            elif path.suffix.lower() == ".xlsx":
                self.xlsx_path.set(str(path))
                self.load_headers(path)
                self.log(f"Excel selected by drag-and-drop: {path}")

    def log(self, text: str):
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")

    def clear_log(self):
        self.log_box.delete("1.0", "end")

    def select_ggf(self):
        path = filedialog.askopenfilename(
            title=self.t("ggf_template"),
            filetypes=[("GeoDIN GGF files", "*.GGF *.ggf"), ("All files", "*.*")]
        )
        if path:
            self.ggf_path.set(path)

            if not self.output_path.get():
                p = Path(path)
                self.output_path.set(str(p.with_name(f"{p.stem}_updated.GGF")))

    def select_xlsx(self):
        path = filedialog.askopenfilename(
            title=self.t("excel_file"),
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if path:
            self.xlsx_path.set(path)
            self.load_headers(Path(path))

    def select_output(self):
        path = filedialog.asksaveasfilename(
            title=self.t("output_file"),
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
            messagebox.showerror(self.t("excel_error"), str(exc))

    def validate_inputs(self) -> bool:
        if not self.ggf_path.get():
            messagebox.showerror("Error", self.t("missing_ggf"))
            return False

        if not self.xlsx_path.get():
            messagebox.showerror("Error", self.t("missing_excel"))
            return False

        if not self.output_path.get():
            messagebox.showerror("Error", self.t("missing_output"))
            return False

        try:
            float(self.filter_min.get())
            float(self.filter_max.get())
        except ValueError:
            messagebox.showerror("Error", self.t("invalid_filter"))
            return False

        dummy = self.dummy_date.get().strip()

        if len(dummy) != 8 or not dummy.isdigit():
            messagebox.showerror("Error", self.t("invalid_dummy"))
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
            language=self.lang.get(),
        )

        try:
            result = update_ggf_dates(settings)

            self.log("\n".join(result.log_lines))
            self.log("")
            self.log(f"{self.t('report_saved')} {report_path}")

            if result.duplicate_dates and not self.remove_duplicates.get():
                messagebox.showwarning(
                    self.t("duplicate_title"),
                    self.t("duplicate_message"),
                )

            messagebox.showinfo(
                self.t("finished"),
                build_user_summary(result) + f"\n{self.t('report_saved')}\n{report_path}",
            )

        except Exception as exc:
            self.log(f"ERROR: {exc}")
            messagebox.showerror(self.t("processing_error"), str(exc))


def run_app():
    if DND_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()

    GeoDINDateToolApp(root)
    root.mainloop()