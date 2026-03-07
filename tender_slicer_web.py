#!/usr/bin/env python3
"""
标书切片工具 - Web 版本
使用浏览器访问 http://localhost:8000
"""

import io
import logging
import zipfile
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

from docx import Document
from docx.document import Document as DocxDocument
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table
from docx.text.paragraph import Paragraph
from flask import Flask, jsonify, render_template_string, request
from werkzeug.exceptions import ClientDisconnected, RequestEntityTooLarge

# =============================================================================
# Flask 应用配置
# =============================================================================

app = Flask(__name__)

UPLOAD_FOLDER = Path(__file__).parent / 'download'
UPLOAD_FOLDER.mkdir(exist_ok=True)

app.config.update({
    'UPLOAD_FOLDER': str(UPLOAD_FOLDER),
    'MAX_CONTENT_LENGTH': 1024 * 1024 * 1024,  # 1GB
})

logging.getLogger('werkzeug').setLevel(logging.ERROR)

# =============================================================================
# 标书切片器类
# =============================================================================

class TenderSlicer:
    """标书切片器 - 将 Word 文档按目录大纲结构切片为多个编号的 Markdown 文件"""

    def __init__(self, docx_path):
        self.docx_path = Path(docx_path)
        self.doc = None
        self.sections = []

    def iter_block_items(self, parent):
        """遍历文档中的所有块元素（段落和表格），保持原始顺序"""
        if hasattr(parent, 'element'):
            parent_elm = parent.element.body
        else:
            parent_elm = parent

        for element in parent_elm.iterchildren():
            if isinstance(element, CT_P):
                yield Paragraph(element, parent)
            elif isinstance(element, CT_Tbl):
                yield Table(element, parent)

    def extract_paragraph_images(self, paragraph):
        """检测段落中的图片，返回占位符标记"""
        images = []
        for run in paragraph.runs:
            for inline in run._element.xpath('.//w:drawing/wp:inline'):
                try:
                    blip = inline.xpath('.//a:blip')
                    if blip:
                        # 检测到图片，记录占位符
                        images.append("<!-- [图片] -->\n")
                except Exception:
                    continue
        return images

    def extract_table_images(self, table):
        """提取表格单元格中的所有图片"""
        images = []
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    images.extend(self.extract_paragraph_images(paragraph))
        return images

    def load_document(self):
        """加载 Word 文档"""
        if not self.docx_path.exists():
            raise FileNotFoundError(f"文件不存在: {self.docx_path}")
        self.doc = Document(str(self.docx_path))

    def get_heading_level(self, paragraph):
        """获取段落的大纲级别，0=正文, 1=一级标题, 2=二级标题, ..."""
        if hasattr(paragraph, '_element'):
            p = paragraph._element
            if p.pPr is not None and hasattr(p.pPr, 'outlineLvl') and p.pPr.outlineLvl is not None:
                return int(p.pPr.outlineLvl.val) + 1
        return 0

    def table_to_markdown(self, table, start_no=1):
        """将表格转换为 Markdown 格式，带编号"""
        if not table.rows:
            return None, start_no

        lines = []
        no = start_no

        header_cells = [cell.text.strip().replace('\n', ' ') for cell in table.rows[0].cells]
        lines.append(f"<!-- {no} --> | " + " | ".join(header_cells) + " |")
        lines.append("<!-- " + str(no + 1) + " --> |" + "|".join(["---"] * len(header_cells)) + "|")
        no += 2

        for row in table.rows[1:]:
            cells = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
            if all(cell in ("", " ") for cell in cells):
                continue
            lines.append(f"<!-- {no} --> | " + " | ".join(cells) + " |")
            no += 1

        return '\n'.join(lines) + '\n\n', no

    def sanitize_filename(self, filename):
        """清理文件名，移除非法字符"""
        import re
        filename = re.sub(r'[<>:"/\\|?*]', '', filename)
        if len(filename) > 100:
            filename = filename[:100]
        return filename.strip()

    def slice_document(self, max_level=None):
        """切片文档，每行添加编号，保留所有表格和图片"""
        self.load_document()

        if max_level is None or max_level == 0:
            max_level = float('inf')

        sections = []
        section_stack = []
        section_index = 0
        line_no = 1

        cover_section = {
            'level': 0,
            'title': '封面',
            'content': [],
            'index': section_index
        }
        section_stack.append(cover_section)

        for block in self.iter_block_items(self.doc):
            if isinstance(block, Paragraph):
                level = self.get_heading_level(block)
                text = block.text.strip()
                images = self.extract_paragraph_images(block)

                if not text and not images:
                    continue

                if any(kw in text for kw in ['目录', '目  录', 'CONTENTS']):
                    continue

                if level > 0:
                    if level <= max_level:
                        if section_stack[-1]['content']:
                            sections.append(section_stack[-1])
                            section_index += 1

                        new_section = {
                            'level': level,
                            'title': text,
                            'content': [],
                            'index': section_index
                        }

                        while section_stack and section_stack[-1]['level'] > level:
                            section_stack.pop()

                        section_stack.append(new_section)
                        section_stack[-1]['content'].append(f"<!-- {line_no} --> {'#' * level} {text}\n")
                        line_no += 1
                    else:
                        section_stack[-1]['content'].append(f"<!-- {line_no} --> {'#' * level} {text}\n")
                        line_no += 1
                else:
                    if text:
                        section_stack[-1]['content'].append(f"<!-- {line_no} --> {text}\n")
                        line_no += 1

                for img in images:
                    section_stack[-1]['content'].append(f"<!-- {line_no} --> {img}")
                    line_no += 1

            elif isinstance(block, Table):
                table_md, line_no = self.table_to_markdown(block, line_no)
                if table_md:
                    section_stack[-1]['content'].append(table_md)

                table_images = self.extract_table_images(block)
                for img in table_images:
                    section_stack[-1]['content'].append(f"<!-- {line_no} --> {img}")
                    line_no += 1

        for section in section_stack:
            if section['content']:
                sections.append(section)

        self.sections = sections
        return sections


# =============================================================================
# 错误处理
# =============================================================================

@app.errorhandler(RequestEntityTooLarge)
def handle_request_entity_too_large(e):
    return jsonify({'error': '文件大小超过限制，最大支持 1 GB'}), 413


@app.errorhandler(ClientDisconnected)
def handle_client_disconnected(e):
    return jsonify({'error': '上传中断'}), 400


# =============================================================================
# 路由定义
# =============================================================================

@app.route('/')
def index():
    return render_template_string(_INDEX_HTML)


@app.route('/slice', methods=['POST'])
def slice_file():
    """切片文件接口，返回 zip 文件"""
    upload_path = None
    try:
        if 'file' not in request.files:
            return jsonify({'error': '未上传文件'}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': '文件名为空'}), 400

        if not file.filename.endswith('.docx'):
            return jsonify({'error': '只支持 .docx 格式'}), 400

        max_level = request.form.get('max_level', '0')
        try:
            if max_level.lower() == 'all':
                max_level = 0
            else:
                max_level = int(max_level)
                if max_level not in [0, 1, 2, 3]:
                    return jsonify({'error': '无效的切分层级'}), 400
        except ValueError:
            return jsonify({'error': '切分层级参数格式错误'}), 400

        upload_path = UPLOAD_FOLDER / file.filename
        file.save(str(upload_path))

        slicer = TenderSlicer(upload_path)
        sections = slicer.slice_document(max_level=max_level if max_level > 0 else None)

        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for section in sections:
                index = section['index'] + 1
                index_str = str(index).zfill(3)
                safe_title = slicer.sanitize_filename(section['title'])
                filename = f"{index_str}_{safe_title}.md"
                content = ''.join(section['content'])
                zipf.writestr(filename, content.encode('utf-8'))

            index_content = "# 标书切片索引\n\n"
            index_content += f"原文件: {file.filename}\n"
            index_content += f"切片时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            index_content += f"总章节数: {len(sections)}\n\n---\n\n## 章节列表\n\n"

            for section in sections:
                idx = section['index'] + 1
                title = section['title']
                filename = f"{str(idx).zfill(3)}_{slicer.sanitize_filename(title)}.md"
                index_content += f"- [{idx}. {title}]({filename})\n"

            zipf.writestr("00_index.md", index_content.encode('utf-8'))

        zip_buffer.seek(0)

        try:
            upload_path.unlink()
        except Exception:
            pass

        safe_filename = quote(f"sliced_{file.filename}.zip", safe='')
        return app.response_class(
            zip_buffer.getvalue(),
            mimetype='application/zip',
            headers={
                'Content-Disposition': f"attachment; filename*=UTF-8''{safe_filename}",
                'X-Section-Count': str(len(sections))
            }
        )

    except Exception as e:
        import traceback
        traceback.print_exc()

        if upload_path and upload_path.exists():
            try:
                upload_path.unlink()
            except Exception:
                pass

        error_msg = str(e)
        if len(error_msg) > 500:
            error_msg = error_msg[:500] + '...'
        return jsonify({'error': error_msg}), 500


# =============================================================================
# HTML 模板
# =============================================================================

_INDEX_HTML = '''<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>标书切片工具</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 20px;
        }
        .container {
            background: white;
            border-radius: 20px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            padding: 40px;
            max-width: 600px;
            width: 100%;
        }
        h1 { color: #333; margin-bottom: 30px; text-align: center; font-size: 28px; }
        .upload-area {
            border: 3px dashed #667eea;
            border-radius: 15px;
            padding: 50px 20px;
            text-align: center;
            background: #f8f9ff;
            transition: all 0.3s;
            cursor: pointer;
        }
        .upload-area:hover, .upload-area.dragover {
            border-color: #764ba2;
            background: #f0f2ff;
        }
        .icon { font-size: 48px; margin-bottom: 15px; }
        .text { color: #666; font-size: 16px; }
        .subtext { color: #999; font-size: 14px; margin-top: 8px; }
        #fileInput { display: none; }
        .btn {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border: none;
            padding: 12px 30px;
            border-radius: 25px;
            font-size: 16px;
            cursor: pointer;
            display: block;
            margin: 20px auto 0;
        }
        .btn:hover { transform: translateY(-2px); box-shadow: 0 5px 20px rgba(102, 126, 234, 0.4); }
        .btn:disabled { background: #ccc; cursor: not-allowed; transform: none; }
        #fileInfo {
            background: #f0f2ff;
            padding: 15px;
            border-radius: 10px;
            margin-top: 20px;
            display: none;
        }
        #fileInfo.show { display: block; }
        #progress {
            margin-top: 20px;
            display: none;
        }
        #progress.show { display: block; }
        .progress-bar {
            background: #e0e0e0;
            border-radius: 10px;
            height: 8px;
            overflow: hidden;
        }
        .progress-fill {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            height: 100%;
            width: 0%;
            transition: width 0.3s;
        }
        #result { margin-top: 20px; display: none; }
        #result.show { display: block; }
        .success {
            background: #f0fff4;
            border: 1px solid #52C41A;
            color: #389e0d;
            padding: 15px;
            border-radius: 10px;
            text-align: center;
        }
        .download-btn {
            background: #52C41A;
            color: white;
            border: none;
            padding: 14px 28px;
            border-radius: 30px;
            font-size: 16px;
            cursor: pointer;
            margin-top: 10px;
            text-decoration: none;
            display: inline-block;
        }
        .download-btn:hover { background: #389e0d; }
        .error {
            background: #fff2f0;
            border: 1px solid #ff4d4f;
            color: #ff4d4f;
            padding: 15px;
            border-radius: 10px;
        }
        #levelSelector {
            margin-top: 20px;
            background: #f8f9ff;
            padding: 20px;
            border-radius: 12px;
        }
        #levelSelector label {
            display: block;
            color: #666;
            font-size: 14px;
            font-weight: 500;
            margin-bottom: 12px;
        }
        .level-options { display: flex; gap: 10px; }
        .level-option { cursor: pointer; position: relative; }
        .level-option input[type="radio"] { display: none; }
        .level-option span {
            display: inline-block;
            padding: 10px 20px;
            background: white;
            border: 2px solid #e0e0e0;
            border-radius: 8px;
            color: #666;
            font-size: 14px;
            transition: all 0.3s;
        }
        .level-option input[type="radio"]:checked + span {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            border-color: #667eea;
            color: white;
        }
        .level-option:hover span { border-color: #667eea; }
    </style>
</head>
<body>
    <div class="container">
        <h1>📄 标书切片工具</h1>

        <div class="upload-area" id="uploadArea">
            <div class="icon">📁</div>
            <div class="text">点击或拖拽上传投标文件</div>
            <div class="subtext">支持 .docx 格式，最大 1GB</div>
        </div>

        <input type="file" id="fileInput" accept=".docx">

        <div id="fileInfo"></div>

        <div id="levelSelector">
            <label>切分层级</label>
            <div class="level-options">
                <label class="level-option">
                    <input type="radio" name="sliceLevel" value="1">
                    <span>一级</span>
                </label>
                <label class="level-option">
                    <input type="radio" name="sliceLevel" value="2" checked>
                    <span>二级</span>
                </label>
                <label class="level-option">
                    <input type="radio" name="sliceLevel" value="3">
                    <span>三级</span>
                </label>
                <label class="level-option">
                    <input type="radio" name="sliceLevel" value="all">
                    <span>全部</span>
                </label>
            </div>
        </div>

        <button class="btn" id="sliceBtn" disabled>开始切片</button>

        <div id="progress">
            <div style="text-align: center; color: #666; margin-bottom: 8px;" id="progressText">准备中...</div>
            <div class="progress-bar">
                <div class="progress-fill" id="progressFill"></div>
            </div>
        </div>

        <div id="result"></div>
    </div>

    <script>
        const uploadArea = document.getElementById('uploadArea');
        const fileInput = document.getElementById('fileInput');
        const fileInfo = document.getElementById('fileInfo');
        const sliceBtn = document.getElementById('sliceBtn');
        const progress = document.getElementById('progress');
        const progressFill = document.getElementById('progressFill');
        const progressText = document.getElementById('progressText');
        const result = document.getElementById('result');

        let selectedFile = null;

        uploadArea.addEventListener('click', () => fileInput.click());

        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.classList.add('dragover');
        });

        uploadArea.addEventListener('dragleave', () => {
            uploadArea.classList.remove('dragover');
        });

        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.classList.remove('dragover');
            if (e.dataTransfer.files.length > 0) {
                handleFile(e.dataTransfer.files[0]);
            }
        });

        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                handleFile(e.target.files[0]);
            }
        });

        function handleFile(file) {
            const MAX_SIZE = 1024 * 1024 * 1024;
            if (file.size > MAX_SIZE) {
                showError('文件大小超过 1GB 限制');
                return;
            }

            if (!file.name.endsWith('.docx')) {
                showError('请上传 .docx 格式的文件');
                return;
            }

            selectedFile = file;
            fileInfo.innerHTML = `✓ 已选择: <strong>${file.name}</strong><br>大小: ${formatFileSize(file.size)}`;
            fileInfo.classList.add('show');
            sliceBtn.disabled = false;
            result.innerHTML = '';
            result.classList.remove('show');
        }

        function formatFileSize(bytes) {
            if (bytes < 1024) return bytes + ' B';
            if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(2) + ' KB';
            if (bytes < 1024 * 1024 * 1024) return (bytes / (1024 * 1024)).toFixed(2) + ' MB';
            return (bytes / (1024 * 1024 * 1024)).toFixed(2) + ' GB';
        }

        function showError(message) {
            result.innerHTML = `<div class="error">${message}</div>`;
            result.classList.add('show');
            sliceBtn.disabled = true;
            fileInfo.classList.remove('show');
        }

        sliceBtn.addEventListener('click', async () => {
            if (!selectedFile) return;

            progress.classList.add('show');
            result.innerHTML = '';
            result.classList.remove('show');
            sliceBtn.disabled = true;

            const formData = new FormData();
            formData.append('file', selectedFile);
            formData.append('max_level', document.querySelector('input[name="sliceLevel"]:checked').value);

            updateProgress(10, '正在上传...');

            const controller = new AbortController();
            const timeoutMs = Math.max(300000, selectedFile.size * 0.001);
            const timeoutId = setTimeout(() => controller.abort(), timeoutMs);

            try {
                const response = await fetch('/slice', {
                    method: 'POST',
                    body: formData,
                    signal: controller.signal
                });

                clearTimeout(timeoutId);
                updateProgress(80, '正在切片...');

                if (!response.ok) {
                    let errorMsg = '处理失败';
                    try {
                        const errorData = await response.json();
                        errorMsg = errorData.error || errorMsg;
                    } catch {}
                    throw new Error(errorMsg);
                }

                updateProgress(100, '处理完成！');

                const blob = await response.blob();
                const url = URL.createObjectURL(blob);

                result.innerHTML = `<div class="success">✅ 切片完成！共 ${response.headers.get('X-Section-Count') || '多个'} 个章节</div>`;

                const downloadBtn = document.createElement('a');
                downloadBtn.href = url;
                downloadBtn.download = 'sliced_documents.zip';
                downloadBtn.className = 'download-btn';
                downloadBtn.textContent = '⬇️ 下载切片结果';
                downloadBtn.style.display = 'block';
                downloadBtn.style.textAlign = 'center';
                result.appendChild(downloadBtn);
                result.classList.add('show');

            } catch (error) {
                clearTimeout(timeoutId);
                showError(`❌ ${error.message}`);
            }

            sliceBtn.disabled = false;
        });

        function updateProgress(value, text) {
            progressFill.style.width = value + '%';
            progressText.textContent = text;
        }
    </script>
</body>
</html>'''


# =============================================================================
# 主程序入口
# =============================================================================

if __name__ == '__main__':
    print("标书切片工具 - Web 版本")
    print("请在浏览器中访问: http://localhost:8000")
    print("按 Ctrl+C 停止服务器")
    app.run(host='0.0.0.0', port=8000, debug=True, threaded=True)
