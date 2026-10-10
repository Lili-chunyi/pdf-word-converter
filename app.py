# -*- coding: utf-8 -*-
"""PDF ⇄ Word 双向转换工具 - 本地 Web 服务

纯本地处理：文件不上传任何云端服务器，仅在浏览器与
本机 127.0.0.1 之间传输。
"""
import json
import os
import re
import shutil
import threading
import time
import uuid
import webbrowser

from flask import Flask, jsonify, make_response, render_template, request, send_file

from converter import pdf2word, word2pdf

APP_VERSION = "1.3.2"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STORAGE_DIR = os.path.join(BASE_DIR, "storage")
RETENTION_DAYS = 30  # 历史记录与结果文件保留时长

MAX_UPLOAD_MB = 200
PDF_EXTS = {".pdf"}
WORD_EXTS = {".docx", ".doc"}

# 客户端隔离：每个浏览器一个随机ID，历史与文件按 ID 分目录互不可见。
# localhost 视为同一台机器，统一使用固定 ID "local"。
_CID_RE = re.compile(r"^(local|[A-Za-z0-9_-]{8,64})$")

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024
app.config["JSON_AS_ASCII"] = False

_tasks = {}          # task_id -> 任务状态字典（含 owner 客户端ID）
_lock = threading.Lock()


def _cid():
    """从请求中提取客户端ID（URL参数 cid 或请求头 X-Client-Id），无效返回 None。"""
    v = request.args.get("cid") or request.headers.get("X-Client-Id") or ""
    return v if _CID_RE.match(v) else None


def _client_dir(cid):
    return os.path.join(STORAGE_DIR, cid)


def _client_tasks_dir(cid):
    return os.path.join(_client_dir(cid), "tasks")


def _history_path(cid):
    return os.path.join(_client_dir(cid), "history.json")


# ---------------------------------------------------------------------------
# 历史记录持久化（storage/<客户端ID>/history.json）
# ---------------------------------------------------------------------------
def _load_history(cid):
    try:
        with open(_history_path(cid), "r", encoding="utf-8") as f:
            data = json.load(f)
            return data if isinstance(data, list) else []
    except Exception:
        return []


def _save_history(cid, records):
    os.makedirs(_client_dir(cid), exist_ok=True)
    tmp = _history_path(cid) + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False)
    os.replace(tmp, _history_path(cid))


def _add_history_record(cid, record):
    with _lock:
        records = _load_history(cid)
        records = [r for r in records if r.get("task_id") != record["task_id"]]
        records.append(record)
        _save_history(cid, records)


def _delete_history_record(cid, task_id):
    with _lock:
        records = _load_history(cid)
        kept = [r for r in records if r.get("task_id") != task_id]
        if len(kept) != len(records):
            _save_history(cid, kept)
            return True
        return False


def _find_task_meta(cid, task_id):
    """先查内存中的活跃任务（校验归属），再查该客户端的持久化历史。"""
    with _lock:
        t = _tasks.get(task_id)
        if t is not None:
            if t.get("owner") != cid:
                return None
            if t["status"] == "done":
                return {"download_name": t["download_name"]}
            return None
    for rec in _load_history(cid):
        if rec.get("task_id") == task_id:
            return rec
    return None


def _cleanup_storage():
    """启动时清理：迁移旧版单文件布局，删除超期(30天)任务目录与记录、孤儿目录。"""
    os.makedirs(STORAGE_DIR, exist_ok=True)
    cutoff = time.time() - RETENTION_DAYS * 86400

    # 旧版布局迁移：storage/history.json + storage/<task_id>/ → storage/local/
    legacy = os.path.join(STORAGE_DIR, "history.json")
    if os.path.exists(legacy):
        local_tasks = os.path.join(STORAGE_DIR, "local", "tasks")
        os.makedirs(local_tasks, exist_ok=True)
        for name in os.listdir(STORAGE_DIR):
            p = os.path.join(STORAGE_DIR, name)
            if os.path.isdir(p) and name != "local":
                dst = os.path.join(local_tasks, name)
                if not os.path.exists(dst):
                    shutil.move(p, dst)
                else:
                    shutil.rmtree(p, ignore_errors=True)
        os.replace(legacy, _history_path("local"))

    for name in os.listdir(STORAGE_DIR):
        cdir = os.path.join(STORAGE_DIR, name)
        if not os.path.isdir(cdir) or not _CID_RE.match(name):
            shutil.rmtree(cdir, ignore_errors=True)  # 清理非法/遗留目录
            continue
        tdir = _client_tasks_dir(name)
        os.makedirs(tdir, exist_ok=True)
        keep = []
        for rec in _load_history(name):
            task_dir = os.path.join(tdir, rec.get("task_id", ""))
            if rec.get("created", 0) < cutoff:
                shutil.rmtree(task_dir, ignore_errors=True)
            else:
                keep.append(rec)
        known = {r["task_id"] for r in keep}
        for d in os.listdir(tdir):
            if d not in known:
                shutil.rmtree(os.path.join(tdir, d), ignore_errors=True)
        _save_history(name, keep)


def _update(task_id, **fields):
    with _lock:
        task = _tasks.get(task_id)
        if task:
            task.update(fields)


def _run_task(task_id, src_path, original_name, target, options, cid):
    out_dir = os.path.join(_client_tasks_dir(cid), task_id)
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
            _add_history_record(cid, {
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
            _add_history_record(cid, {
                "task_id": task_id, "original_name": original_name,
                "download_name": f"{stem}.pdf", "target": "word2pdf",
                "created": time.time(),
                "size": os.path.getsize(final),
            })
    except Exception as e:
        # 只向用户返回必要信息，抹掉内部路径，避免泄露服务器目录结构
        msg = str(e)
        for secret in (src_path, out_dir, STORAGE_DIR):
            if secret:
                msg = msg.replace(secret, "文件")
        if isinstance(e, (pdf2word.ConversionError, word2pdf.ConversionError)):
            _update(task_id, status="error", message=msg)
        else:
            print(f"[task {task_id}] 转换异常: {msg}")
            _update(task_id, status="error",
                    message="转换失败，请确认文件未损坏后重试；若反复失败请换一个文件试试")


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
    cid = _cid()
    if not cid:
        return jsonify({"error": "缺少客户端标识，请刷新页面重试"}), 403
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
    task_dir = os.path.join(_client_tasks_dir(cid), task_id)
    os.makedirs(task_dir, exist_ok=True)
    src_path = os.path.join(task_dir, f"input{ext}")
    f.save(src_path)

    with _lock:
        _tasks[task_id] = {
            "status": "processing", "progress": 2,
            "message": "文件已接收，正在准备转换…",
            "download_name": None, "created": time.time(),
            "owner": cid,
        }

    threading.Thread(
        target=_run_task,
        args=(task_id, src_path, original_name, target, options, cid),
        daemon=True).start()
    return jsonify({"task_id": task_id})


@app.get("/api/status/<task_id>")
def api_status(task_id):
    cid = _cid()
    with _lock:
        task = _tasks.get(task_id)
        if not task or not cid or task.get("owner") != cid:
            return jsonify({"error": "任务不存在"}), 404
        return jsonify({k: task[k] for k in
                        ("status", "progress", "message", "download_name")})


@app.get("/api/history")
def api_history():
    """该浏览器近 30 天的转换记录（仅含成功任务），按时间倒序。"""
    cid = _cid()
    if not cid:
        return jsonify({"error": "缺少客户端标识，请刷新页面重试"}), 403
    cutoff = time.time() - RETENTION_DAYS * 86400
    records = [r for r in _load_history(cid) if r.get("created", 0) >= cutoff]
    records.sort(key=lambda r: r.get("created", 0), reverse=True)
    tdir = _client_tasks_dir(cid)
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
            "available": os.path.exists(os.path.join(tdir, r.get("task_id", ""),
                                                     f"output{ext}")),
        })
    return jsonify(out)


@app.delete("/api/history/<task_id>")
def api_history_delete(task_id):
    """删除该浏览器的一条历史记录及其结果文件。"""
    cid = _cid()
    if not cid:
        return jsonify({"error": "缺少客户端标识，请刷新页面重试"}), 403
    if not _delete_history_record(cid, task_id):
        return jsonify({"error": "记录不存在"}), 404
    with _lock:
        t = _tasks.get(task_id)
        if t is not None and t.get("owner") == cid:
            _tasks.pop(task_id, None)
    shutil.rmtree(os.path.join(_client_tasks_dir(cid), task_id),
                  ignore_errors=True)
    return jsonify({"ok": True})


@app.delete("/api/history")
def api_history_clear():
    """清空该浏览器的全部历史记录及结果文件。"""
    cid = _cid()
    if not cid:
        return jsonify({"error": "缺少客户端标识，请刷新页面重试"}), 403
    for rec in _load_history(cid):
        tid = rec.get("task_id", "")
        shutil.rmtree(os.path.join(_client_tasks_dir(cid), tid),
                      ignore_errors=True)
        with _lock:
            t = _tasks.get(tid)
            if t is not None and t.get("owner") == cid:
                _tasks.pop(tid, None)
    _save_history(cid, [])
    return jsonify({"ok": True})


@app.get("/api/download/<task_id>")
def api_download(task_id):
    cid = _cid()
    if not cid:
        return jsonify({"error": "缺少客户端标识"}), 403
    meta = _find_task_meta(cid, task_id)
    if not meta:
        return jsonify({"error": "任务不存在或尚未完成"}), 404
    ext = ".docx" if meta["download_name"].lower().endswith(".docx") else ".pdf"
    path = os.path.join(_client_tasks_dir(cid), task_id, f"output{ext}")
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
    cid = _cid()
    if not cid:
        return jsonify({"error": "缺少客户端标识"}), 403
    meta = _find_task_meta(cid, task_id)
    if not meta:
        return jsonify({"error": "任务不存在或尚未完成"}), 404

    out_dir = os.path.join(_client_tasks_dir(cid), task_id)
    if meta["download_name"].lower().endswith(".pdf"):
        preview_path = os.path.join(out_dir, "output.pdf")
        if not os.path.exists(preview_path):
            return jsonify({"error": "结果文件已过期，请重新转换"}), 404
    else:
        preview_path = os.path.join(out_dir, "output.pdf")
        if not os.path.exists(preview_path) or os.path.getsize(preview_path) == 0:
            with _preview_lock(task_id):
                if not os.path.exists(preview_path) \
                        or os.path.getsize(preview_path) == 0:
                    try:
                        word2pdf.convert_word_to_pdf(
                            os.path.join(out_dir, "output.docx"), out_dir)
                    except Exception as e:
                        msg = str(e).replace(out_dir, "文件").replace(
                            STORAGE_DIR, "storage")
                        return jsonify({"error": f"预览生成失败：{msg}"}), 500

    return send_file(preview_path, mimetype="application/pdf",
                     as_attachment=False)


@app.errorhandler(413)
def too_large(e):
    return jsonify({"error": f"文件超过 {MAX_UPLOAD_MB}MB 上限"}), 413


if __name__ == "__main__":
    _cleanup_storage()
    host = os.environ.get("HOST", "127.0.0.1")
    port = int(os.environ.get("PORT", "8765"))
    url = f"http://{'127.0.0.1' if host in ('0.0.0.0', '::') else host}:{port}/"
    if port == 8765:
        print("=" * 52)
        print("  PDF ⇄ Word 双向转换工具已启动（纯本地处理）")
        print(f"  请在浏览器访问：{url}")
        print("  关闭本窗口即退出工具")
        print("=" * 52)
    if not os.environ.get("PDFWORD_NO_BROWSER"):
        threading.Timer(1.2, lambda: webbrowser.open(url)).start()
    app.run(host=host, port=port, threaded=True, debug=False)
