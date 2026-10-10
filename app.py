# -*- coding: utf-8 -*-
"""PDF ⇄ Word 双向转换工具 - 本地 Web 服务

纯本地处理：文件不上传任何云端服务器，仅在浏览器与
本机 127.0.0.1 之间传输。
"""
import json
import os
import shutil
import threading
import time
import uuid
import webbrowser

from flask import Flask, jsonify, make_response, render_template, request, send_file

from converter import pdf2word, word2pdf

APP_VERSION = "1.2.4"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STORAGE_DIR = os.path.join(BASE_DIR, "storage")
HISTORY_FILE = os.path.join(STORAGE_DIR, "history.json")
RETENTION_DAYS = 30  # 历史记录与结果文件保留时长

MAX_UPLOAD_MB = 200
PDF_EXTS = {".pdf"}
WORD_EXTS = {".docx", ".doc"}

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024
app.config["JSON_AS_ASCII"] = False

_tasks = {}          # task_id -> 任务状态字典
_lock = threading.Lock()


# ---------------------------------------------------------------------------
# 历史记录持久化（storage/history.json）
# ---------------------------------------------------------------------------
def _load_history():
    try:
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data if isinstance(data, list) else []
    except Exception:
        return []


def _save_history(records):
    tmp = HISTORY_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False)
    os.replace(tmp, HISTORY_FILE)


def _add_history_record(record):
    with _lock:
        records = _load_history()
        records = [r for r in records if r.get("task_id") != record["task_id"]]
        records.append(record)
        _save_history(records)


def _delete_history_record(task_id):
    with _lock:
        records = _load_history()
        kept = [r for r in records if r.get("task_id") != task_id]
        if len(kept) != len(records):
            _save_history(kept)
            return True
        return False


def _find_task_meta(task_id):
    """先查内存中的活跃任务，再查持久化历史记录。"""
    with _lock:
        t = _tasks.get(task_id)
        if t is not None:
            if t["status"] == "done":
                return {"download_name": t["download_name"]}
            return None
    for rec in _load_history():
        if rec.get("task_id") == task_id:
            return rec
    return None


def _cleanup_storage():
    """启动时清理：删除超期(30天)任务目录与记录、以及无记录的孤儿目录。"""
    os.makedirs(STORAGE_DIR, exist_ok=True)
    cutoff = time.time() - RETENTION_DAYS * 86400
    keep = []
    for rec in _load_history():
        task_dir = os.path.join(STORAGE_DIR, rec.get("task_id", ""))
        if rec.get("created", 0) < cutoff:
            shutil.rmtree(task_dir, ignore_errors=True)
        else:
            keep.append(rec)
    known = {r["task_id"] for r in keep}
    for name in os.listdir(STORAGE_DIR):
        if name == "history.json":
            continue
        p = os.path.join(STORAGE_DIR, name)
        if os.path.isdir(p) and name not in known:
            shutil.rmtree(p, ignore_errors=True)
    _save_history(keep)


def _update(task_id, **fields):
    with _lock:
        task = _tasks.get(task_id)
        if task:
            task.update(fields)


def _run_task(task_id, src_path, original_name, target, options):
    out_dir = os.path.join(STORAGE_DIR, task_id)
    stem = os.path.splitext(original_name)[0]
    try:
        if target == "pdf2word":
            out_path = os.path.join(out_dir, "output.docx")
            used, pages = pdf2word.convert_pdf_to_word(
                src_path, out_path,
                mode=options.get("mode", "auto"),
                dpi=int(options.get("dpi", 200)),
                progress_cb=lambda p, m: _update(task_id, progress=p, message=m),
            )
            label = "可编辑文本版" if used == "text" else "图片版"
            _update(task_id, status="done", progress=100,
                    message=f"转换成功（{label}，共 {pages} 页）",
                    download_name=f"{stem}.docx")
            _add_history_record({
                "task_id": task_id, "original_name": original_name,
                "download_name": f"{stem}.docx", "target": "pdf2word",
                "created": time.time(),
                "size": os.path.getsize(out_path),
            })
        else:
            out_path = word2pdf.convert_word_to_pdf(
                src_path, out_dir,
                progress_cb=lambda p, m: _update(task_id, progress=p, message=m))
            final = os.path.join(out_dir, "output.pdf")
            if os.path.abspath(out_path) != os.path.abspath(final):
                shutil.move(out_path, final)
            _update(task_id, status="done", progress=100,
                    message="转换成功",
                    download_name=f"{stem}.pdf")
            _add_history_record({
                "task_id": task_id, "original_name": original_name,
                "download_name": f"{stem}.pdf", "target": "word2pdf",
                "created": time.time(),
                "size": os.path.getsize(final),
            })
    except Exception as e:
        _update(task_id, status="error", message=str(e))


@app.get("/")
def index():
    resp = make_response(render_template("index.html", version=APP_VERSION))
    # 禁用缓存，避免浏览器停留在旧版页面
    resp.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    return resp


@app.get("/api/version")
def api_version():
    return jsonify({"version": APP_VERSION})


@app.post("/api/convert")
def api_convert():
    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify({"error": "请先选择文件"}), 400

    original_name = os.path.basename(f.filename)
    ext = os.path.splitext(original_name)[1].lower()
    target = request.form.get("target", "")

    if ext in PDF_EXTS:
        expected = "pdf2word"
    elif ext in WORD_EXTS:
        expected = "word2pdf"
    else:
        return jsonify({"error": "仅支持 PDF / DOCX / DOC 文件"}), 400
    if target != expected:
        return jsonify({"error": "文件类型与转换方向不匹配"}), 400

    options = {
        "mode": request.form.get("mode", "auto"),
        "dpi": request.form.get("dpi", "200"),
    }

    task_id = uuid.uuid4().hex[:12]
    task_dir = os.path.join(STORAGE_DIR, task_id)
    os.makedirs(task_dir, exist_ok=True)
    src_path = os.path.join(task_dir, f"input{ext}")
    f.save(src_path)

    with _lock:
        _tasks[task_id] = {
            "status": "processing", "progress": 2,
            "message": "文件已接收，正在准备转换…",
            "download_name": None, "created": time.time(),
        }

    threading.Thread(
        target=_run_task, args=(task_id, src_path, original_name, target, options),
        daemon=True).start()
    return jsonify({"task_id": task_id})


@app.get("/api/status/<task_id>")
def api_status(task_id):
    with _lock:
        task = _tasks.get(task_id)
        if not task:
            return jsonify({"error": "任务不存在"}), 404
        return jsonify({k: task[k] for k in
                        ("status", "progress", "message", "download_name")})


@app.get("/api/history")
def api_history():
    """近 30 天的转换记录（仅含成功任务），按时间倒序。"""
    cutoff = time.time() - RETENTION_DAYS * 86400
    records = [r for r in _load_history() if r.get("created", 0) >= cutoff]
    records.sort(key=lambda r: r.get("created", 0), reverse=True)
    out = []
    for r in records:
        ext = ".docx" if r.get("download_name", "").lower().endswith(".docx") \
            else ".pdf"
        out.append({
            "task_id": r.get("task_id"),
            "original_name": r.get("original_name", ""),
            "download_name": r.get("download_name", ""),
            "target": r.get("target", ""),
            "created": r.get("created", 0),
            "size": r.get("size", 0),
            "available": os.path.exists(
                os.path.join(STORAGE_DIR, r.get("task_id", ""), f"output{ext}")),
        })
    return jsonify(out)


@app.delete("/api/history/<task_id>")
def api_history_delete(task_id):
    """删除单条历史记录及其结果文件。"""
    if not _delete_history_record(task_id):
        return jsonify({"error": "记录不存在"}), 404
    with _lock:
        _tasks.pop(task_id, None)
    shutil.rmtree(os.path.join(STORAGE_DIR, task_id), ignore_errors=True)
    return jsonify({"ok": True})


@app.delete("/api/history")
def api_history_clear():
    """清空全部历史记录及结果文件。"""
    for rec in _load_history():
        tid = rec.get("task_id", "")
        shutil.rmtree(os.path.join(STORAGE_DIR, tid), ignore_errors=True)
        with _lock:
            _tasks.pop(tid, None)
    _save_history([])
    return jsonify({"ok": True})


@app.get("/api/download/<task_id>")
def api_download(task_id):
    meta = _find_task_meta(task_id)
    if not meta:
        return jsonify({"error": "任务不存在或尚未完成"}), 404
    ext = ".docx" if meta["download_name"].lower().endswith(".docx") else ".pdf"
    path = os.path.join(STORAGE_DIR, task_id, f"output{ext}")
    if not os.path.exists(path):
        return jsonify({"error": "结果文件已过期或已被清理，请重新转换"}), 404
    return send_file(path, as_attachment=True,
                     download_name=meta["download_name"])


_preview_locks = {}


def _preview_lock(task_id):
    with _lock:
        return _preview_locks.setdefault(task_id, threading.Lock())


@app.get("/api/preview/<task_id>")
def api_preview(task_id):
    """内联返回结果 PDF 供页面预览。

    - Word→PDF 任务：直接返回 output.pdf
    - PDF→Word 任务：docx 无法被浏览器直接渲染，用 LibreOffice
      转成预览 PDF 并缓存到任务目录，首次稍慢，之后即时打开。
    """
    meta = _find_task_meta(task_id)
    if not meta:
        return jsonify({"error": "任务不存在或尚未完成"}), 404

    if meta["download_name"].lower().endswith(".pdf"):
        preview_path = os.path.join(STORAGE_DIR, task_id, "output.pdf")
        if not os.path.exists(preview_path):
            return jsonify({"error": "结果文件已过期，请重新转换"}), 404
    else:
        out_dir = os.path.join(STORAGE_DIR, task_id)
        preview_path = os.path.join(out_dir, "output.pdf")
        if not os.path.exists(preview_path) or os.path.getsize(preview_path) == 0:
            with _preview_lock(task_id):
                if not os.path.exists(preview_path) \
                        or os.path.getsize(preview_path) == 0:
                    try:
                        word2pdf.convert_word_to_pdf(
                            os.path.join(out_dir, "output.docx"), out_dir)
                    except Exception as e:
                        return jsonify({"error": f"预览生成失败：{e}"}), 500

    return send_file(preview_path, mimetype="application/pdf",
                     as_attachment=False)


@app.errorhandler(413)
def too_large(e):
    return jsonify({"error": f"文件超过 {MAX_UPLOAD_MB}MB 上限"}), 413


if __name__ == "__main__":
    _cleanup_storage()
    url = "http://127.0.0.1:8765/"
    print("=" * 52)
    print("  PDF ⇄ Word 双向转换工具已启动（纯本地处理）")
    print(f"  请在浏览器访问：{url}")
    print("  关闭本窗口即退出工具")
    print("=" * 52)
    if not os.environ.get("PDFWORD_NO_BROWSER"):
        threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    app.run(host="127.0.0.1", port=8765, threaded=True, debug=False)
