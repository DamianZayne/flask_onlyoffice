"""
智慧办案助手 — Flask 应用
支持 OnlyOffice 离线文档编辑
"""

import json
import os
import re
import uuid
from datetime import datetime
from pathlib import Path

import requests
from docx import Document
from flask import (
    Flask,
    abort,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    send_from_directory,
    url_for,
)

# ── 配置 ──────────────────────────────────────────────────────
BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)
META_FILE = BASE_DIR / "meta.json"

MIME_MAP = {
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    "pdf":  "application/pdf",
    "doc":  "application/msword",
    "xls":  "application/vnd.ms-excel",
    "ppt":  "application/vnd.ms-powerpoint",
    "odt":  "application/vnd.oasis.opendocument.text",
    "ods":  "application/vnd.oasis.opendocument.spreadsheet",
    "odp":  "application/vnd.oasis.opendocument.presentation",
    "csv":  "text/csv",
    "txt":  "text/plain",
}

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024

# ── 材料名 → Edit_Doc 目录映射 ────────────────────────────────
EDIT_DOC_DIR = BASE_DIR / "Edit_Doc"
MATERIAL_DOC_MAP = {
    "安全首课记录": 1,
    "安全责任书": 12,
    "承诺书": 2,
    "谈话方案和安全预案呈阅件": 3,
    "走读式谈话安全评估情况及安全预案": 4,
    "走读式谈话安全评估表": 13,
    "谈话方案": 14,
    "采取谈话（询问、讯问）措施呈批表": 5,
    "谈话对象人员名单": 15,
    "暂予保管物品登记表": 7,
    "谈话对象体检情况登记表": 6,
    "陪送交接单": 16,
    "谈话（讯问、询问）工作记录": 8,
    "有关工作情况说明": 17,
    "审查调查安全责任明细表": 9,  # 同一目录
    "专章总结记录": 11,
    "谈话录音录像交接表": 10,
}


# ── 元数据 ────────────────────────────────────────────────────
def load_meta():
    if META_FILE.exists():
        return json.loads(META_FILE.read_text(encoding="utf-8"))
    return {}

def save_meta(meta):
    META_FILE.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")


# ── SDK 静态资源 ──────────────────────────────────────────────
@app.route("/packages/<path:filename>")
def serve_packages(filename):
    return send_from_directory(BASE_DIR / "static" / "packages", filename)


# ── 业务页面 ──────────────────────────────────────────────────

@app.route("/")
def index():
    """首页 - 工作台"""
    return render_template("首页.html")


@app.route("/login")
def login():
    """登录页"""
    return render_template("登录.html")


@app.route("/security-catalog")
def security_catalog():
    """安全卷录入"""
    return render_template("安全卷录入.html")


# ── OnlyOffice 编辑器 ──────────────    ───────────────────────────

@app.route("/editor/<doc_id>")
def edit_doc(doc_id):
    """编辑已上传的文档"""
    meta = load_meta()
    info = meta.get(doc_id)
    if not info:
        abort(404)

    filepath = UPLOAD_DIR / info["filename"]
    if not filepath.exists():
        abort(404)

    return render_template(
        "editor.html",
        doc_id=doc_id,
        doc_name=info["name"],
        doc_ext=info["ext"],
        doc_url=url_for("download_doc", doc_id=doc_id, _external=True),
        save_url=url_for("save_doc", doc_id=doc_id, _external=True),
        back_url="/",
    )


@app.route("/editor/material")
def edit_material():
    """编辑安全卷材料文档 —— 优先加载 Edit_Doc 里预置的模板"""
    material_name = request.args.get("name", "未命名材料")
    case_id = request.args.get("id", "new")

    # 查找 Edit_Doc 中对应的文档
    dir_num = MATERIAL_DOC_MAP.get(material_name)
    doc_url = ""

    doc_ext = "docx"
    doc_name = f"{material_name}.docx"

    if dir_num:
        doc_dir = EDIT_DOC_DIR / str(dir_num)
        if doc_dir.exists():
            files = list(doc_dir.iterdir())
            if files:
                doc_file = files[0]  # 取第一个文件
                doc_ext = doc_file.suffix.lstrip(".").lower()
                doc_name = doc_file.name
                doc_url = url_for("serve_edit_doc", dir_num=dir_num,
                                  filename=doc_file.name, _external=True)

    save_url = ""
    if dir_num and doc_url:
        save_url = url_for("save_edit_doc", dir_num=dir_num,
                           filename=doc_name, _external=True)

    return render_template(
        "editor.html",
        doc_id=f"material_{case_id}_{dir_num or 'new'}",
        doc_name=doc_name,
        doc_ext=doc_ext,
        doc_url=doc_url,
        save_url=save_url,
        back_url="/security-catalog",
    )


@app.route("/edit-doc/<int:dir_num>/<path:filename>")
def serve_edit_doc(dir_num, filename):
    """提供 Edit_Doc 目录下的文档"""
    filepath = EDIT_DOC_DIR / str(dir_num) / filename
    if not filepath.exists():
        abort(404)
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else "docx"
    return send_file(filepath, as_attachment=False,
                     mimetype=MIME_MAP.get(ext, "application/octet-stream"))


# ── 文档 API ─────────────────────────────────────────────────

@app.route("/api/doc/<doc_id>/download")
def download_doc(doc_id):
    """下载已上传的文档"""
    meta = load_meta()
    info = meta.get(doc_id)
    if not info:
        abort(404)

    filepath = UPLOAD_DIR / info["filename"]
    if not filepath.exists():
        abort(404)

    return send_file(filepath, as_attachment=False, mimetype=MIME_MAP.get(info["ext"], "application/octet-stream"))


@app.route("/api/doc/<doc_id>/info")
def doc_info(doc_id):
    meta = load_meta()
    info = meta.get(doc_id)
    if not info:
        abort(404)
    return jsonify(info)


# ── 保存接口 ─────────────────────────────────────────────────

@app.route("/api/save-doc/<doc_id>", methods=["POST"])
def save_doc(doc_id):
    """保存已上传文档的修改"""
    meta = load_meta()
    info = meta.get(doc_id)
    if not info:
        abort(404)

    filepath = UPLOAD_DIR / info["filename"]
    filepath.write_bytes(request.get_data())
    print(f"[save] 已保存: {filepath}")
    return jsonify({"error": 0})


@app.route("/api/save-edit-doc/<int:dir_num>/<path:filename>", methods=["POST"])
def save_edit_doc(dir_num, filename):
    """保存 Edit_Doc 材料文档的修改 — 覆写目录下已有文件"""
    doc_dir = EDIT_DOC_DIR / str(dir_num)
    # 找到目录中已有的文件，覆写它（不创建新文件）
    existing = list(doc_dir.glob("*")) if doc_dir.exists() else []
    target = existing[0] if existing else doc_dir / filename
    target.write_bytes(request.get_data())
    print(f"[save-edit-doc] 已保存: {target}")
    return jsonify({"error": 0})


# ── 批量替换 ─────────────────────────────────────────────────

@app.route("/api/edit-doc-list")
def edit_doc_list():
    """列出 Edit_Doc 中所有可供替换的文档（仅 .docx，python-docx 不支持 .doc）"""
    files = []
    if EDIT_DOC_DIR.exists():
        for d in sorted(EDIT_DOC_DIR.iterdir(), key=lambda x: int(x.name) if x.name.isdigit() else 999):
            if d.is_dir():
                for f in d.iterdir():
                    if f.suffix.lower() == ".docx":
                        material_name = ""
                        for name, num in MATERIAL_DOC_MAP.items():
                            if num == int(d.name):
                                material_name = name
                                break
                        files.append({
                            "dir_num": int(d.name),
                            "filename": f.name,
                            "material": material_name,
                            "path": f"{d.name}/{f.name}",
                        })
    return jsonify(files)


@app.route("/api/batch-replace", methods=["POST"])
def batch_replace():
    """批量替换 Edit_Doc 中文档的文本，保留格式"""
    data = request.get_json(force=True)
    rel_path = data.get("file_path", "")
    replacements = data.get("replacements", {})

    if not rel_path or not replacements:
        return jsonify({"error": 1, "message": "缺少参数"}), 400

    filepath = EDIT_DOC_DIR / rel_path
    if not filepath.exists():
        return jsonify({"error": 1, "message": f"文件不存在: {rel_path}"}), 404

    if filepath.suffix.lower() != ".docx":
        return jsonify({"error": 1, "message": "仅支持 .docx 格式，.doc 请先用 Word 另存为 .docx"}), 400

    try:
        doc = Document(str(filepath))

        def replace_in_runs(runs, reps):
            for run in runs:
                for old_text, new_text in reps.items():
                    if old_text in run.text:
                        run.text = run.text.replace(old_text, new_text)

        # 段落
        for paragraph in doc.paragraphs:
            replace_in_runs(paragraph.runs, replacements)

        # 表格
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        replace_in_runs(paragraph.runs, replacements)

        # 页眉页脚
        for section in doc.sections:
            for paragraph in section.header.paragraphs:
                replace_in_runs(paragraph.runs, replacements)
            for paragraph in section.footer.paragraphs:
                replace_in_runs(paragraph.runs, replacements)

        # 保存回原文件
        doc.save(str(filepath))

        count = len(replacements)
        print(f"[batch-replace] 已替换 {count} 组词 → {filepath}")
        return jsonify({"error": 0, "message": f"成功替换 {count} 组文本", "count": count})

    except Exception as e:
        print(f"[batch-replace] 失败: {e}")
        return jsonify({"error": 1, "message": str(e)}), 500


# ── 启动 ──────────────────────────────────────────────────────
if __name__ == "__main__":
    import sys
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 5050
    print("=" * 60)
    print("  智慧办案助手 - 离线版")
    print(f"  访问: http://localhost:{port}")
    print("=" * 60)
    app.run(host="0.0.0.0", port=port, debug=True)
