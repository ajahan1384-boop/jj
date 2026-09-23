#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
رابط وب محلی برای «به‌روزرسانی قیمت اکسل دیجی‌کالا».

این یک سرور وب است که فقط روی خود کامپیوتر شما اجرا می‌شود (127.0.0.1) و
هیچ فایلی به بیرون از سیستم شما ارسال نمی‌کند. صفحه‌ی HTML در static/index.html
رابط کاربری را نشان می‌دهد؛ پردازش واقعی فایل‌ها همان منطق core.py است که
در پردازه‌های جداگانه (ProcessPoolExecutor) و با تعداد محدود اجرا می‌شود تا
فشاری به سیستم وارد نشود.

اجرا:
    pip install -r requirements.txt
    python web_app.py

سپس مرورگر به‌صورت خودکار باز می‌شود (یا خودتان به آدرس چاپ‌شده بروید).
"""

from __future__ import annotations

import io
import os
import re
import sys
import time
import uuid
import shutil
import threading
import webbrowser
import zipfile
from concurrent.futures import ProcessPoolExecutor, as_completed

from flask import Flask, jsonify, request, send_file, send_from_directory, abort

from core import process_one_file, MAX_RECOMMENDED_FILES

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STATIC_DIR = os.path.join(BASE_DIR, "static")
JOBS_DIR = os.path.join(BASE_DIR, "_jobs")
HOST = "127.0.0.1"
PORT = int(os.environ.get("PORT", "8765"))

app = Flask(__name__, static_folder=None)

_jobs_lock = threading.Lock()
JOBS: dict[str, dict] = {}


def _safe_filename(name: str) -> str:
    """نام فایل را از کاراکترهای خطرناک/مسیر پاک می‌کند ولی متن فارسی را حفظ می‌کند."""
    name = os.path.basename((name or "").strip()) or "file.xlsx"
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name)
    name = name.strip(" .") or "file"
    if not name.lower().endswith((".xlsx", ".xls")):
        name += ".xlsx"
    return name


def _new_job(total: int) -> str:
    job_id = uuid.uuid4().hex
    with _jobs_lock:
        JOBS[job_id] = {
            "total": total,
            "done": 0,
            "success": 0,
            "fail": 0,
            "results": [],
            "finished": False,
            "cancelled": False,
            "executor": None,
        }
    return job_id


def _job_dirs(job_id: str):
    uploads = os.path.join(JOBS_DIR, job_id, "uploads")
    outputs = os.path.join(JOBS_DIR, job_id, "outputs")
    os.makedirs(uploads, exist_ok=True)
    os.makedirs(outputs, exist_ok=True)
    return uploads, outputs


def _run_job(job_id: str, saved_files: list[str], multiplier_str: str, workers: int):
    _, outputs = _job_dirs(job_id)
    try:
        executor = ProcessPoolExecutor(max_workers=workers)
        with _jobs_lock:
            JOBS[job_id]["executor"] = executor
        try:
            futures = {
                executor.submit(process_one_file, path, outputs, multiplier_str): path
                for path in saved_files
            }
            for future in as_completed(futures):
                path = futures[future]
                try:
                    result = future.result()
                except Exception as exc:  # noqa: BLE001
                    from core import FileResult
                    result = FileResult(path, False, f"خطای پردازش: {exc}")
                with _jobs_lock:
                    job = JOBS[job_id]
                    job["done"] += 1
                    entry = {
                        "filename": os.path.basename(result.input_path),
                        "ok": result.ok,
                        "message": result.message,
                        "download_name": os.path.basename(result.output_path) if result.ok else None,
                        "output_path": result.output_path if result.ok else None,
                    }
                    job["results"].append(entry)
                    if result.ok:
                        job["success"] += 1
                    else:
                        job["fail"] += 1
        finally:
            executor.shutdown(wait=True, cancel_futures=True)
    except Exception as exc:  # noqa: BLE001 -- شبکه ایمنی نهایی برای کل دسته
        with _jobs_lock:
            job = JOBS[job_id]
            job["results"].append({
                "filename": "-", "ok": False,
                "message": f"خطای کلی در اجرای دسته‌ای: {exc}",
                "download_name": None, "output_path": None,
            })
            job["fail"] += 1
    finally:
        with _jobs_lock:
            JOBS[job_id]["finished"] = True
            JOBS[job_id]["executor"] = None


@app.route("/")
def index():
    return send_from_directory(STATIC_DIR, "index.html")


@app.route("/api/system")
def system_info():
    cpu = os.cpu_count() or 2
    return jsonify({
        "cpu_count": cpu,
        "default_workers": max(1, cpu - 1),
        "max_recommended_files": MAX_RECOMMENDED_FILES,
    })


@app.route("/api/process", methods=["POST"])
def api_process():
    try:
        multiplier_str = (request.form.get("multiplier") or "").strip()
        if not multiplier_str:
            return jsonify({"error": "ضریب را وارد کنید."}), 400
        try:
            from decimal import Decimal, InvalidOperation
            Decimal(multiplier_str)
        except InvalidOperation:
            return jsonify({"error": "ضریب واردشده یک عدد معتبر نیست."}), 400

        files = request.files.getlist("files")
        files = [f for f in files if f and f.filename]
        if not files:
            return jsonify({"error": "حداقل یک فایل اکسل انتخاب کنید."}), 400

        cpu = os.cpu_count() or 2
        try:
            workers = int(request.form.get("workers") or 0)
        except ValueError:
            workers = 0
        if workers <= 0:
            workers = max(1, cpu - 1)
        workers = max(1, min(workers, cpu))

        job_id = _new_job(len(files))
        uploads, _ = _job_dirs(job_id)

        saved_files = []
        for idx, f in enumerate(files):
            safe_name = _safe_filename(f.filename)
            folder = os.path.join(uploads, f"f{idx}")
            os.makedirs(folder, exist_ok=True)
            path = os.path.join(folder, safe_name)
            f.save(path)
            saved_files.append(path)

        thread = threading.Thread(
            target=_run_job, args=(job_id, saved_files, multiplier_str, workers), daemon=True
        )
        thread.start()

        return jsonify({"job_id": job_id, "total": len(files)})
    except Exception as exc:  # noqa: BLE001
        return jsonify({"error": f"خطای غیرمنتظره سرور: {exc}"}), 500


@app.route("/api/jobs/<job_id>")
def api_job_status(job_id):
    with _jobs_lock:
        job = JOBS.get(job_id)
        if job is None:
            return jsonify({"error": "چنین پردازشی پیدا نشد."}), 404
        results = list(job["results"])
        payload = {
            "total": job["total"],
            "done": job["done"],
            "success": job["success"],
            "fail": job["fail"],
            "finished": job["finished"],
            "results": [
                {
                    "filename": r["filename"],
                    "ok": r["ok"],
                    "message": r["message"],
                    "download_url": (
                        f"/api/jobs/{job_id}/download/{i}" if r["ok"] else None
                    ),
                }
                for i, r in enumerate(results)
            ],
        }
    return jsonify(payload)


@app.route("/api/jobs/<job_id>/cancel", methods=["POST"])
def api_job_cancel(job_id):
    with _jobs_lock:
        job = JOBS.get(job_id)
        if job is None:
            return jsonify({"error": "چنین پردازشی پیدا نشد."}), 404
        executor = job.get("executor")
        job["cancelled"] = True
    if executor is not None:
        try:
            executor.shutdown(wait=False, cancel_futures=True)
        except TypeError:
            executor.shutdown(wait=False)
    return jsonify({"ok": True})


@app.route("/api/jobs/<job_id>/download/<int:idx>")
def api_download_one(job_id, idx):
    with _jobs_lock:
        job = JOBS.get(job_id)
        if job is None:
            abort(404)
        results = job["results"]
        if idx < 0 or idx >= len(results):
            abort(404)
        entry = results[idx]
        path = entry.get("output_path")
    if not entry["ok"] or not path or not os.path.isfile(path):
        abort(404)
    return send_file(path, as_attachment=True, download_name=entry["download_name"])


@app.route("/api/jobs/<job_id>/download-all")
def api_download_all(job_id):
    with _jobs_lock:
        job = JOBS.get(job_id)
        if job is None:
            abort(404)
        ok_entries = [r for r in job["results"] if r["ok"] and r.get("output_path")]
    if not ok_entries:
        abort(404)

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        used_names = set()
        for entry in ok_entries:
            path = entry["output_path"]
            if not os.path.isfile(path):
                continue
            arcname = entry["download_name"] or os.path.basename(path)
            base, ext = os.path.splitext(arcname)
            n = 2
            while arcname in used_names:
                arcname = f"{base}_{n}{ext}"
                n += 1
            used_names.add(arcname)
            zf.write(path, arcname)
    buffer.seek(0)
    return send_file(
        buffer, as_attachment=True, download_name="digikala-price-updater-output.zip",
        mimetype="application/zip",
    )


def _cleanup_old_jobs(max_age_seconds: int = 24 * 3600):
    if not os.path.isdir(JOBS_DIR):
        return
    now = time.time()
    for name in os.listdir(JOBS_DIR):
        path = os.path.join(JOBS_DIR, name)
        try:
            if os.path.isdir(path) and (now - os.path.getmtime(path)) > max_age_seconds:
                shutil.rmtree(path, ignore_errors=True)
        except OSError:
            pass


def _open_browser_later():
    time.sleep(1.0)
    try:
        webbrowser.open(f"http://{HOST}:{PORT}/")
    except Exception:  # noqa: BLE001
        pass


def main():
    if sys.platform.startswith("win"):
        import multiprocessing
        multiprocessing.freeze_support()

    os.makedirs(JOBS_DIR, exist_ok=True)
    _cleanup_old_jobs()

    threading.Thread(target=_open_browser_later, daemon=True).start()

    print(f"سرور محلی در حال اجراست: http://{HOST}:{PORT}/")
    print("برای توقف، همین پنجره ترمینال را ببندید یا Ctrl+C بزنید.")
    try:
        app.run(host=HOST, port=PORT, threaded=True, debug=False)
    except OSError as exc:
        print(f"خطا در اجرای سرور روی پورت {PORT}: {exc}")
        print("احتمالاً پورت در حال استفاده است؛ برنامه دیگری را ببندید یا متغیر محیطی PORT را تغییر دهید.")


if __name__ == "__main__":
    main()
