# 🍉 PDF ⇄ Word 双向转换工具

一个轻量、安全、易用的本地文档格式转换工具。打开网页即可上传、转换、下载，**无需安装 Office 或 Adobe**。

## 功能特性

### PDF → Word
- **文本型 PDF**：转为可编辑 Word，保留文字、字体、字号、颜色、表格、图片、超链接等
- **扫描件 / 图片型 PDF**：自动识别后按原顺序、原方向、原比例逐页嵌入高清页面图片（可选 150/200/300 DPI），保证完整清晰、无空白页
- 自动校验页数与图片数一致，异常时自动回退图片版

### Word → PDF
- 基于 LibreOffice 引擎稳定转换，尽量保持原排版，无需安装 Office

### 界面功能
- 拖拽上传，方向自动识别，也可手动指定
- 转换完成直接在页面下方**预览结果**（Word 结果自动渲染预览 PDF）
- 右侧「历史记录」：近 30 天转换记录持久化保存，可预览、下载、单条删除、一键清空
- 页面自动更新：服务升级后打开的页面自动刷新

## 隐私承诺

**所有转换均在本地完成，任何文件都不会上传到网络。** 历史记录与结果文件只保存在本机 `storage/` 目录（该目录不会进入 git 仓库），30 天后自动清理。

## 环境要求

- Python 3.9+
- [LibreOffice](https://www.libreoffice.org/)（仅 Word → PDF 需要；Windows 可用 `winget install TheDocumentFoundation.LibreOffice` 安装）

## 安装与运行

```bash
pip install -r requirements.txt
python app.py        # 或 Windows 下双击 启动工具.bat
```

浏览器会自动打开 `http://127.0.0.1:8765`。

## 项目结构

```
├── app.py                  # Flask 服务（上传/进度/下载/预览/历史）
├── converter/
│   ├── pdf2word.py         # PDF→Word 双引擎（文本版 / 图片版）
│   └── word2pdf.py         # Word→PDF（LibreOffice）
├── templates/index.html    # 前端页面
├── tests/                  # 测试脚本与样例生成器
├── storage/                # 运行时数据（git 忽略，勿提交）
└── 启动工具.bat             # Windows 一键启动
```
