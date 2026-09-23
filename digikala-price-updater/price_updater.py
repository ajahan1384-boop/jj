#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
به‌روزرسانی دسته‌ای ستون قیمت فروش (ستون J) در فایل‌های اکسل دیجی‌کالا.

هر عدد ستون J در شیت «داده ها» در ضریب وارد شده ضرب می‌شود و نتیجه تا ۴ رقم
اعشار به سمت بالا (Ceiling) گرد می‌شود. بقیه‌ی فایل (شیت راهنما، استایل‌ها،
ستون‌های دیگر) دقیقاً دست‌نخورده باقی می‌ماند چون فایل ورودی به‌جای ساخت از
نو، مستقیماً ویرایش و ذخیره می‌شود.

اجرا:
    pip install -r requirements.txt
    python price_updater.py
"""

from __future__ import annotations

import os
import sys
import queue
import threading
from decimal import Decimal, InvalidOperation
from concurrent.futures import ProcessPoolExecutor, as_completed

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

from core import FileResult, process_one_file, MAX_RECOMMENDED_FILES

try:
    import openpyxl
except ImportError:
    openpyxl = None


# --------------------------------------------------------------------------
# رابط کاربری
# --------------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("به‌روزرسانی قیمت اکسل دیجی‌کالا")
        self.geometry("760x600")
        self.minsize(680, 520)

        self.selected_files: list[str] = []
        self.output_dir: str = ""
        self.result_queue: "queue.Queue" = queue.Queue()
        self.running = False
        self.executor: ProcessPoolExecutor | None = None
        self.futures = []
        self.total_files = 0
        self.done_files = 0
        self.success_count = 0
        self.fail_count = 0

        self._build_ui()
        self._poll_queue()

    # ---------------- UI construction ----------------
    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        top = ttk.Frame(self)
        top.pack(fill="x", **pad)

        ttk.Label(top, text="ضریب قیمت:").grid(row=0, column=0, sticky="e")
        self.multiplier_var = tk.StringVar(value="1")
        ttk.Entry(top, textvariable=self.multiplier_var, width=12, justify="center").grid(
            row=0, column=1, sticky="w", padx=(4, 20)
        )

        default_workers = max(1, (os.cpu_count() or 2) - 1)
        max_workers = os.cpu_count() or 4
        ttk.Label(top, text="پردازش هم‌زمان:").grid(row=0, column=2, sticky="e")
        self.workers_var = tk.IntVar(value=default_workers)
        ttk.Spinbox(
            top, from_=1, to=max_workers, textvariable=self.workers_var, width=5, justify="center"
        ).grid(row=0, column=3, sticky="w", padx=4)
        ttk.Label(top, text=f"(حداکثر توصیه‌شده: {max_workers})", foreground="#666").grid(
            row=0, column=4, sticky="w", padx=(4, 0)
        )

        files_frame = ttk.LabelFrame(self, text="فایل‌های ورودی (حداکثر پیشنهادی ۵۰ فایل هم‌زمان)")
        files_frame.pack(fill="both", expand=False, **pad)

        btns = ttk.Frame(files_frame)
        btns.pack(fill="x", padx=8, pady=6)
        ttk.Button(btns, text="انتخاب فایل‌های اکسل...", command=self.choose_files).pack(side="right")
        ttk.Button(btns, text="پاک کردن لیست", command=self.clear_files).pack(side="right", padx=6)
        self.files_count_label = ttk.Label(btns, text="۰ فایل انتخاب شده")
        self.files_count_label.pack(side="left")

        self.files_list = tk.Listbox(files_frame, height=6)
        self.files_list.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        out_frame = ttk.LabelFrame(self, text="پوشه خروجی")
        out_frame.pack(fill="x", **pad)
        out_row = ttk.Frame(out_frame)
        out_row.pack(fill="x", padx=8, pady=6)
        self.output_dir_var = tk.StringVar(value="(به‌صورت پیش‌فرض کنار فایل‌های ورودی، پوشه output ساخته می‌شود)")
        ttk.Label(out_row, textvariable=self.output_dir_var, foreground="#444").pack(side="right", fill="x", expand=True)
        ttk.Button(out_row, text="انتخاب پوشه...", command=self.choose_output_dir).pack(side="left")

        run_frame = ttk.Frame(self)
        run_frame.pack(fill="x", **pad)
        self.start_btn = ttk.Button(run_frame, text="شروع پردازش", command=self.start_processing)
        self.start_btn.pack(side="right")
        self.cancel_btn = ttk.Button(run_frame, text="لغو", command=self.cancel_processing, state="disabled")
        self.cancel_btn.pack(side="right", padx=6)

        self.progress = ttk.Progressbar(self, mode="determinate")
        self.progress.pack(fill="x", padx=10, pady=(0, 6))

        self.status_var = tk.StringVar(value="آماده")
        ttk.Label(self, textvariable=self.status_var).pack(fill="x", padx=10)

        log_frame = ttk.LabelFrame(self, text="گزارش")
        log_frame.pack(fill="both", expand=True, **pad)
        self.log = scrolledtext.ScrolledText(log_frame, height=12, state="disabled")
        self.log.pack(fill="both", expand=True, padx=6, pady=6)

    # ---------------- actions ----------------
    def choose_files(self):
        paths = filedialog.askopenfilenames(
            title="انتخاب فایل‌های اکسل",
            filetypes=[("Excel files", "*.xlsx"), ("همه فایل‌ها", "*.*")],
        )
        if not paths:
            return
        self.selected_files = list(paths)
        if len(self.selected_files) > MAX_RECOMMENDED_FILES:
            messagebox.showwarning(
                "تعداد فایل زیاد",
                f"{len(self.selected_files)} فایل انتخاب شده. حداکثر توصیه‌شده {MAX_RECOMMENDED_FILES} فایل هم‌زمان است؛ "
                "پردازش همه‌ی آن‌ها انجام می‌شود ولی ممکن است بیشتر طول بکشد.",
            )
        self._refresh_files_list()

    def clear_files(self):
        self.selected_files = []
        self._refresh_files_list()

    def _refresh_files_list(self):
        self.files_list.delete(0, tk.END)
        for p in self.selected_files:
            self.files_list.insert(tk.END, os.path.basename(p))
        self.files_count_label.config(text=f"{len(self.selected_files)} فایل انتخاب شده")

    def choose_output_dir(self):
        d = filedialog.askdirectory(title="انتخاب پوشه خروجی")
        if d:
            self.output_dir = d
            self.output_dir_var.set(d)

    def _log(self, text: str):
        self.log.config(state="normal")
        self.log.insert(tk.END, text + "\n")
        self.log.see(tk.END)
        self.log.config(state="disabled")

    def start_processing(self):
        if self.running:
            return
        if not self.selected_files:
            messagebox.showerror("خطا", "ابتدا حداقل یک فایل اکسل انتخاب کنید.")
            return

        multiplier_str = self.multiplier_var.get().strip()
        try:
            Decimal(multiplier_str)
        except InvalidOperation:
            messagebox.showerror("خطا", "ضریب واردشده معتبر نیست. یک عدد مثل 1.15 وارد کنید.")
            return

        if openpyxl is None:
            messagebox.showerror(
                "کتابخانه موجود نیست",
                "کتابخانه openpyxl نصب نیست. دستور زیر را در ترمینال اجرا کنید:\n\npip install openpyxl",
            )
            return

        workers = max(1, min(self.workers_var.get(), os.cpu_count() or 1))

        self.log.config(state="normal")
        self.log.delete("1.0", tk.END)
        self.log.config(state="disabled")

        self.total_files = len(self.selected_files)
        self.done_files = 0
        self.success_count = 0
        self.fail_count = 0
        self.progress.config(maximum=self.total_files, value=0)
        self.status_var.set(f"در حال پردازش... 0/{self.total_files}")

        self.running = True
        self.start_btn.config(state="disabled")
        self.cancel_btn.config(state="normal")

        worker_thread = threading.Thread(
            target=self._run_batch, args=(multiplier_str, workers), daemon=True
        )
        worker_thread.start()

    def cancel_processing(self):
        if not self.running or self.executor is None:
            return
        self.status_var.set("در حال لغو...")
        try:
            self.executor.shutdown(wait=False, cancel_futures=True)
        except TypeError:
            self.executor.shutdown(wait=False)

    def _run_batch(self, multiplier_str: str, workers: int):
        default_out_base = os.path.dirname(self.selected_files[0]) if self.selected_files else "."
        out_dir = self.output_dir or os.path.join(default_out_base, "output")

        try:
            with ProcessPoolExecutor(max_workers=workers) as executor:
                self.executor = executor
                futures = {
                    executor.submit(process_one_file, path, out_dir, multiplier_str): path
                    for path in self.selected_files
                }
                for future in as_completed(futures):
                    path = futures[future]
                    try:
                        result = future.result()
                    except Exception as exc:  # noqa: BLE001
                        result = FileResult(path, False, f"خطای پردازش: {exc}")
                    self.result_queue.put(result)
        except Exception as exc:  # noqa: BLE001
            self.result_queue.put(FileResult("", False, f"خطای کلی در اجرای دسته‌ای: {exc}"))
        finally:
            self.executor = None
            self.result_queue.put(None)  # نشانه پایان

    def _poll_queue(self):
        try:
            while True:
                item = self.result_queue.get_nowait()
                if item is None:
                    self._finish_batch()
                    continue
                self._handle_result(item)
        except queue.Empty:
            pass
        self.after(150, self._poll_queue)

    def _handle_result(self, result: FileResult):
        self.done_files += 1
        self.progress.config(value=self.done_files)
        name = os.path.basename(result.input_path) if result.input_path else "?"
        if result.ok:
            self.success_count += 1
            self._log(f"✅ {name}: {result.message}")
        else:
            self.fail_count += 1
            self._log(f"❌ {name}: {result.message}")
        self.status_var.set(f"در حال پردازش... {self.done_files}/{self.total_files}")

    def _finish_batch(self):
        if not self.running:
            return
        self.running = False
        self.start_btn.config(state="normal")
        self.cancel_btn.config(state="disabled")
        self.status_var.set(
            f"پایان یافت: {self.success_count} موفق، {self.fail_count} ناموفق از {self.total_files} فایل"
        )
        self._log("---")
        self._log(f"پایان پردازش. {self.success_count} موفق، {self.fail_count} ناموفق.")
        if self.fail_count == 0 and self.success_count > 0:
            messagebox.showinfo("پایان", f"همه {self.success_count} فایل با موفقیت پردازش شدند.")
        elif self.success_count > 0:
            messagebox.showwarning(
                "پایان با خطا در برخی فایل‌ها",
                f"{self.success_count} فایل موفق و {self.fail_count} فایل ناموفق بود. جزئیات در گزارش.",
            )
        else:
            messagebox.showerror("پردازش ناموفق", "هیچ فایلی با موفقیت پردازش نشد. جزئیات در گزارش را ببینید.")


def main():
    if sys.platform.startswith("win"):
        # جلوگیری از اجرای مجدد ناخواسته اسکریپت در هر پردازه فرعی روی ویندوز
        import multiprocessing
        multiprocessing.freeze_support()
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
