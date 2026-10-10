# -*- coding: utf-8 -*-
"""客户端隔离安全性测试：模拟两个浏览器 (A/B)，验证历史与文件互不可见。

运行：python -B tests/test_isolation.py
"""
import io
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.chdir(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import app  # noqa: E402

A, B = "aaaaaaaaaaaa", "bbbbbbbbbbbb"
CHECKS = []


def check(name, cond):
    CHECKS.append((name, bool(cond)))
    print(("PASS " if cond else "FAIL ") + name)


# 1. 旧版布局迁移（storage/history.json 旧数据应迁入 storage/local/）
app._cleanup_storage()
check("旧版历史已迁移到 storage/local",
      not os.path.exists(os.path.join("storage", "history.json"))
      and os.path.exists(os.path.join("storage", "local", "history.json")))

c = app.app.test_client()

# 2. 未带/非法 cid 的请求被拒绝
check("无 cid 请求历史被拒(403)", c.get("/api/history").status_code == 403)
check("无 cid 请求转换被拒(403)",
      c.post("/api/convert", data={"target": "pdf2word"}).status_code == 403)
check("非法 cid 被拒(403)", c.get("/api/history?cid=../x").status_code == 403)

# 3. 浏览器 A 转换一个 PDF
import fitz  # noqa: E402

os.makedirs(os.path.join("tests", "out"), exist_ok=True)
pdf_path = os.path.join("tests", "out", "iso.pdf")
doc = fitz.open()
page = doc.new_page()
page.insert_text((72, 72), "isolation test 你好")
doc.save(pdf_path)
doc.close()

with open(pdf_path, "rb") as fh:
    r = c.post("/api/convert?cid=" + A,
               data={"file": (fh, "iso.pdf"), "target": "pdf2word"})
check("浏览器A 上传转换成功", r.status_code == 200)
tid = r.get_json()["task_id"]

for _ in range(120):
    st = c.get(f"/api/status/{tid}?cid=" + A).get_json()
    if st.get("status") in ("done", "error"):
        break
    time.sleep(0.5)
check("浏览器A 转换完成", st.get("status") == "done")

# 4. 隔离验证
hist_b = c.get("/api/history?cid=" + B).get_json()
check("浏览器B 历史为空(看不到A的记录)", hist_b == [])
hist_a = c.get("/api/history?cid=" + A).get_json()
check("浏览器A 历史包含自己的记录",
      len(hist_a) == 1 and hist_a[0]["task_id"] == tid)
check("浏览器B 下载A的文件被拒(404)",
      c.get(f"/api/download/{tid}?cid=" + B).status_code == 404)
check("浏览器A 下载自己的文件成功(200)",
      c.get(f"/api/download/{tid}?cid=" + A).status_code == 200)
check("浏览器B 查询A的任务状态被拒(404)",
      c.get(f"/api/status/{tid}?cid=" + B).status_code == 404)
check("浏览器B 预览A的文件被拒(404)",
      c.get(f"/api/preview/{tid}?cid=" + B).status_code == 404)
check("浏览器B 删除A的记录被拒(404)",
      c.delete(f"/api/history/{tid}?cid=" + B).status_code == 404)
check("浏览器B 的一键清空不影响A",
      (c.delete("/api/history?cid=" + B).status_code == 200)
      and len(c.get("/api/history?cid=" + A).get_json()) == 1)

# 5. A 自己的删除操作正常
check("浏览器A 删除单条记录成功",
      c.delete(f"/api/history/{tid}?cid=" + A).status_code == 200)
check("删除后A历史为空", c.get("/api/history?cid=" + A).get_json() == [])
check("删除后磁盘任务目录同步清除",
      not os.path.exists(os.path.join("storage", A, "tasks", tid)))

# 清理测试产生的客户端目录（不影响 local/）
import shutil  # noqa: E402

for cid in (A, B):
    shutil.rmtree(os.path.join("storage", cid), ignore_errors=True)

passed = sum(1 for _, v in CHECKS if v)
print(f"\n{passed}/{len(CHECKS)} 通过")
sys.exit(0 if passed == len(CHECKS) else 1)
