# -*- coding: utf-8 -*-
"""Word → PDF 转换模块

引擎：LibreOffice 无头模式（soffice --headless）。
无需安装 Microsoft Office / Adobe，由本工具自带转换能力。
支持 .docx 与 .doc 输入，输出尽量保持原排版。

Windows 注意：soffice.exe 仅为启动器，实际转换由后台 soffice.bin
完成，因此不能依赖进程退出判断结束，必须轮询输出文件。
"""
import glob
import os
import shutil
import subprocess
import tempfile
import threading
import time

# 串行化转换：LibreOffice 同一用户配置目录不支持并发
_soffice_lock = threading.Lock()
PROFILE_DIR = None


class ConversionError(Exception):
    """转换失败，message 面向最终用户展示。"""


def find_soffice():
    """定位 LibreOffice 的 soffice 可执行文件。"""
    p = shutil.which("soffice")
    if p:
        return p
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/usr/bin/soffice",
        "/usr/local/bin/soffice",
        "/opt/libreoffice/program/soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    # 兜底：各盘符默认安装路径下再找一次
    for drive in ("C", "D", "E"):
        for c in glob.glob(rf"{drive}:\Program Files*\LibreOffice*\program\soffice.exe"):
            return c
    return None


def _user_profile(fresh=False):
    """独立用户配置目录，避免与本机其他 LibreOffice 实例冲突。"""
    global PROFILE_DIR
    if fresh or PROFILE_DIR is None or not os.path.isdir(PROFILE_DIR):
        PROFILE_DIR = tempfile.mkdtemp(prefix="pdfword_lo_profile_")
    return PROFILE_DIR


def _wait_for_file(path, timeout):
    """轮询等待输出文件生成且大小稳定（soffice.bin 后台写入）。"""
    deadline = time.time() + timeout
    last_size = -1
    while time.time() < deadline:
        if os.path.exists(path):
            size = os.path.getsize(path)
            if size > 0 and size == last_size:
                return True
            last_size = size
        time.sleep(0.4)
    return False


def _run_soffice(cmd, timeout):
    """执行 soffice 命令，超时杀掉整个进程树。"""
    proc = subprocess.Popen(
        cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    try:
        out, err = proc.communicate(timeout=timeout)
        return out, err
    except subprocess.TimeoutExpired:
        try:
            subprocess.run(["taskkill", "/F", "/T", "/PID", str(proc.pid)],
                           capture_output=True)
        except Exception:
            pass
        raise ConversionError("转换超时，文档可能过于复杂，请尝试拆分后重试。")


def convert_word_to_pdf(docx_path, out_dir, progress_cb=None, timeout=300):
    """将 Word 转换为 PDF，返回生成的 PDF 路径。"""
    def report(p, m):
        if progress_cb:
            progress_cb(p, m)

    soffice = find_soffice()
    if not soffice:
        raise ConversionError(
            "未找到 LibreOffice 转换引擎。请安装 LibreOffice（免费开源）后重试，"
            "下载地址：https://www.libreoffice.org/download/download/")

    os.makedirs(out_dir, exist_ok=True)
    base_name = os.path.splitext(os.path.basename(docx_path))[0]
    expected = os.path.join(out_dir, base_name + ".pdf")

    with _soffice_lock:
        report(30, "正在启动转换引擎…")
        last_err = ""
        for attempt, fresh in enumerate((False, True)):
            # 清理上一次尝试可能产生的半成品
            if os.path.exists(expected):
                os.remove(expected)
            profile_uri = "file:///" + _user_profile(fresh=fresh and attempt > 0) \
                .replace("\\", "/")
            cmd = [
                soffice,
                "--headless", "--norestore", "--nolockcheck", "--nodefault",
                f"-env:UserInstallation={profile_uri}",
                "--convert-to", "pdf:writer_pdf_Export",
                "--outdir", out_dir,
                docx_path,
            ]
            out, err = _run_soffice(cmd, timeout)
            last_err = (err or out or "").strip()
            # soffice.bin 可能在 soffice.exe 退出后仍在写入，轮询等待
            if _wait_for_file(expected, 180):
                report(92, "转换完成。")
                return expected
        detail = last_err[-300:] if last_err else "未知原因"
        raise ConversionError(f"Word 转 PDF 失败：{detail}")
