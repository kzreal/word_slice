#!/usr/bin/env python3
"""
标书切片工具 - Web 版本
使用浏览器访问 http://localhost:8000
"""

from flask import Flask, render_template, render_template_string, request, jsonify, send_from_directory
from pathlib import Path
import zipfile
import io
from datetime import datetime
import re
from urllib.parse import quote
import base64
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.document import Document as DocxDocument
from werkzeug.exceptions import RequestEntityTooLarge, ClientDisconnected

try:
    from docx import Document
except ImportError:
    print("错误: 请先安装 python-docx 和 flask")
    print("运行: pip install python-docx flask")
    import sys
    sys.exit(1)

app = Flask(__name__)
# 使用项目目录下的 download 文件夹
UPLOAD_FOLDER = Path(__file__).parent / 'download'
UPLOAD_FOLDER.mkdir(exist_ok=True)
app.config['UPLOAD_FOLDER'] = str(UPLOAD_FOLDER)

# 设置最大上传文件大小限制（2GB）
app.config['MAX_CONTENT_LENGTH'] = 2048 * 1024 * 1024

# 增加 Flask 配置来处理大文件上传
app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 3600  # 1小时
app.config['MAX_CONTENT_LENGTH'] = 2 * 1024 * 1024 * 1024  # 2GB
app.config['UPLOAD_FOLDER'] = str(UPLOAD_FOLDER)

# 增加 Werkzeug 的大文件上传配置
from werkzeug.serving import WSGIRequestHandler
# 禁用请求大小限制日志（避免大文件上传时的性能问题）
import logging
logging.getLogger('werkzeug').setLevel(logging.ERROR)

# 请求日志中间件 - 记录所有请求的详细信息
@app.before_request
def log_request_info():
    """记录所有请求的详细信息"""
    import time
    request.start_time = time.time()
    print(f"[REQUEST] {request.method} {request.path}")
    print(f"[REQUEST] Remote: {request.remote_addr}")
    if request.content_length:
        print(f"[REQUEST] Content-Length: {request.content_length}")
    print(f"[REQUEST] Content-Type: {request.content_type}")
    print(f"[REQUEST] User-Agent: {request.headers.get('User-Agent', 'Unknown')}")
    print(f"[REQUEST] Origin: {request.headers.get('Origin', 'None')}")

@app.after_request
def log_response_info(response):
    """记录所有响应的详细信息"""
    import time
    duration = time.time() - request.start_time
    print(f"[RESPONSE] {request.method} {request.path} - Status: {response.status_code} - Time: {duration:.3f}s")
    return response

# 处理文件过大错误
@app.errorhandler(RequestEntityTooLarge)
def handle_request_entity_too_large(e):
    max_size_mb = app.config['MAX_CONTENT_LENGTH'] // (1024 * 1024)
    return jsonify({
        'error': f'文件大小超过限制，最大支持 {max_size_mb} MB (2GB)'
    }), 413


@app.errorhandler(ClientDisconnected)
def handle_client_disconnected(e):
    print(f"[ERROR] 客户端断开连接: {e}")
    return jsonify({
        'error': '上传中断，可能是网络不稳定或文件太大，请使用桌面版工具'
    }), 400


class TenderSlicer:
    """标书切片器"""

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
        """提取段落中的所有图片，返回 Markdown 图片标记列表"""
        images = []

        # 使用 python-docx 的方式遍历段落中的 run
        for run in paragraph.runs:
            for inline in run._element.xpath('.//w:drawing/wp:inline'):
                try:
                    # 获取图片的 blip
                    blip = inline.xpath('.//a:blip')
                    if blip:
                        embed_id = blip[0].get('{http://schemas.openxmlformats.org/drawingml/2006/main}embed')
                        if embed_id:
                            image_part = self.doc.part.related_parts[embed_id]
                            image_data = image_part.blob
                            image_b64 = base64.b64encode(image_data).decode('utf-8')
                            images.append(f"![图片](data:image/png;base64,{image_b64})\n")
                except Exception:
                    continue

        return images

    def extract_table_images(self, table):
        """提取表格单元格中的所有图片，返回 Markdown 图片标记列表"""
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
        """获取标题级别 - 按大纲层级"""
        if hasattr(paragraph, '_element'):
            p = paragraph._element
            if p.pPr and p.pPr.outlineLvl is not None:
                return int(p.pPr.outlineLvl.val) + 1
        return 0

    def table_to_markdown(self, table):
        """表格转 Markdown"""
        if not table.rows:
            return None
        lines = []
        for i, row in enumerate(table.rows):
            cells = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
            lines.append('| ' + ' | '.join(cells) + ' |')
            if i == 0:
                lines.append('|' + '|'.join(['---'] * len(cells)) + '|')
        return '\n'.join(lines) + '\n\n'

    def sanitize_filename(self, filename):
        """清理文件名"""
        filename = re.sub(r'[<>:"/\\|?*]', '', filename)
        if len(filename) > 100:
            filename = filename[:100]
        return filename.strip()

    def table_to_markdown(self, table, start_no=1):
        """表格转 Markdown，带编号"""
        if not table.rows:
            return None, start_no

        lines = []
        no = start_no

        # 表头
        header_cells = [cell.text.strip().replace('\n', ' ') for cell in table.rows[0].cells]
        lines.append(f"<!-- {no} --> | " + " | ".join(header_cells) + " |")
        lines.append("<!-- " + str(no + 1) + " --> |" + "|".join(["---"] * len(header_cells)) + "|")
        no += 2

        # 数据行
        for row in table.rows[1:]:
            cells = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
            # 跳过空行
            if all(cell in ("", " ") for cell in cells):
                continue
            lines.append(f"<!-- {no} --> | " + " | ".join(cells) + " |")
            no += 1

        return '\n'.join(lines) + '\n\n', no

    def slice_document(self, max_level=None):
        """
        切片文档，每行添加编号，保留所有表格和图片

        Args:
            max_level: 最大切分层级，None或0表示全部层级，1/2/3表示最多切到该层级
        """
        self.load_document()

        # 如果max_level为0或None，设置为极大值表示全部层级
        if max_level is None or max_level == 0:
            max_level = float('inf')

        sections = []
        # 使用栈跟踪当前激活的章节
        section_stack = []
        section_index = 0

        # 初始化封面章节
        cover_section = {
            'level': 0,
            'title': '封面',
            'content': [],
            'index': section_index
        }
        section_stack.append(cover_section)

        line_no = 1  # 全局行号

        # 按文档顺序遍历所有块元素（段落和表格）
        for block in self.iter_block_items(self.doc):
            if isinstance(block, Paragraph):
                # 处理段落
                level = self.get_heading_level(block)
                text = block.text.strip()

                # 提取段落中的图片
                images = self.extract_paragraph_images(block)

                if not text and not images:
                    continue

                if any(kw in text for kw in ['目录', '目  录', 'CONTENTS']):
                    continue

                if level > 0:
                    # 这是一个标题
                    if level <= max_level:
                        # 标题级别在切分范围内，创建新章节

                        # 保存当前章节（如果有内容）
                        if section_stack[-1]['content']:
                            sections.append(section_stack[-1])
                            section_index += 1

                        # 创建新章节
                        new_section = {
                            'level': level,
                            'title': text,
                            'content': [],
                            'index': section_index
                        }

                        # 维护栈：弹出所有级别大于当前级别的章节（注意是 > 不是 >=）
                        while section_stack and section_stack[-1]['level'] > level:
                            section_stack.pop()

                        # 压入新章节
                        section_stack.append(new_section)

                        # 添加带编号的标题到内容
                        section_stack[-1]['content'].append(f"<!-- {line_no} --> {'#' * level} {text}\n")
                        line_no += 1
                    else:
                        # 标题级别超过max_level，作为正文添加到当前章节（不弹出栈）
                        section_stack[-1]['content'].append(f"<!-- {line_no} --> {'#' * level} {text}\n")
                        line_no += 1
                else:
                    # 这是正文，添加到当前章节
                    if text:
                        section_stack[-1]['content'].append(f"<!-- {line_no} --> {text}\n")
                        line_no += 1

                # 添加段落中的图片
                for img in images:
                    section_stack[-1]['content'].append(f"<!-- {line_no} --> {img}")
                    line_no += 1

            elif isinstance(block, Table):
                # 处理表格
                table_md, line_no = self.table_to_markdown(block, line_no)
                if table_md:
                    section_stack[-1]['content'].append(table_md)

                # 提取表格中的图片
                table_images = self.extract_table_images(block)
                for img in table_images:
                    section_stack[-1]['content'].append(f"<!-- {line_no} --> {img}")
                    line_no += 1

        # 保存栈中所有章节
        for section in section_stack:
            if section['content']:
                sections.append(section)

        self.sections = sections
        return sections


@app.route('/')
def index():
    """主页"""
    return render_template_string('''
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>标书切片工具</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
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
        h1 {
            color: #333;
            margin-bottom: 30px;
            font-size: 28px;
            text-align: center;
        }
        .upload-area {
            border: 3px dashed #667eea;
            border-radius: 15px;
            padding: 50px 20px;
            text-align: center;
            background: #f8f9ff;
            transition: all 0.3s;
            cursor: pointer;
        }
        .upload-area:hover {
            border-color: #764ba2;
            background: #f0f2ff;
        }
        .upload-area.dragover {
            border-color: #52C41A;
            background: #f0fff4;
        }
        .icon {
            font-size: 48px;
            margin-bottom: 15px;
        }
        .text {
            color: #666;
            font-size: 16px;
        }
        .subtext {
            color: #999;
            font-size: 14px;
            margin-top: 8px;
        }
        #fileInput {
            display: none;
        }
        .btn {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border: none;
            padding: 12px 30px;
            border-radius: 25px;
            font-size: 16px;
            cursor: pointer;
            transition: all 0.3s;
            display: block;
            margin: 20px auto 0;
        }
        .btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 5px 20px rgba(102, 126, 234, 0.4);
        }
        .btn:disabled {
            background: #ccc;
            cursor: not-allowed;
            transform: none;
        }
        #fileInfo {
            background: #f0f2ff;
            padding: 15px;
            border-radius: 10px;
            margin-top: 20px;
            display: none;
        }
        #fileInfo.show {
            display: block;
        }
        #progress {
            margin-top: 20px;
            display: none;
        }
        #progress.show {
            display: block;
        }
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
        #result {
            margin-top: 20px;
            display: none;
        }
        #result.show {
            display: block;
        }
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
            transition: all 0.3s;
            box-shadow: 0 4px 15px rgba(82, 196, 26, 0.3);
            text-decoration: none;
            display: inline-flex;
            align-items: center;
            gap: 8px;
        }
        .download-btn:hover {
            background: #389e0d;
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(82, 196, 26, 0.4);
        }
        .download-btn-fixed {
            position: fixed;
            bottom: 30px;
            right: 30px;
            z-index: 1000;
            animation: slideIn 0.3s ease-out;
        }
        @keyframes slideIn {
            from {
                opacity: 0;
                transform: translateY(20px);
            }
            to {
                opacity: 1;
                transform: translateY(0);
            }
        }
        .error {
            background: #fff2f0;
            border: 1px solid #ff4d4f;
            color: #ff4d4f;
            padding: 15px;
            border-radius: 10px;
        }
        /* 层级选择器样式 */
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
        .level-options {
            display: flex;
            gap: 10px;
            flex-wrap: wrap;
        }
        .level-option {
            cursor: pointer;
            position: relative;
        }
        .level-option input[type="radio"] {
            display: none;
        }
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
            box-shadow: 0 4px 12px rgba(102, 126, 234, 0.3);
        }
        .level-option:hover span {
            border-color: #667eea;
        }
        #levelHint {
            margin-top: 12px;
            padding: 10px 15px;
            background: #fff7e6;
            border-left: 3px solid #faad14;
            border-radius: 4px;
            color: #856404;
            font-size: 13px;
            line-height: 1.5;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>📄 标书切片工具</h1>

        <div style="background: #fff7e6; border-left: 3px solid #faad14; padding: 12px; border-radius: 4px; margin-bottom: 20px; font-size: 13px; color: #856404; line-height: 1.5;">
            <strong>💡 提示:</strong> 本工具适合处理中小型文件（<50MB）。<br>
            大文件（>100MB）可能因网络问题上传失败。如遇问题，请使用桌面版工具：
            <code style="background: #f0f0f0; padding: 2px 6px; border-radius: 3px;">python tender_slicer_gui.py</code>
        </div>

        <div class="upload-area" id="uploadArea">
            <div class="icon">📁</div>
            <div class="text">点击或拖拽上传投标文件</div>
            <div class="subtext">支持 .docx 格式</div>
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
            <div id="levelHint">最多切到二级标题，三级及以下合并到二级章节</div>
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

        // 页面加载时检查服务器连接
        async function checkServerHealth() {
            try {
                console.log('[HEALTH] 检查服务器连接...');
                const response = await fetch('/health', {
                    method: 'GET',
                    cache: 'no-cache'
                });
                if (response.ok) {
                    const data = await response.json();
                    console.log('[HEALTH] 服务器正常:', data);
                    return true;
                } else {
                    console.warn('[HEALTH] 服务器响应异常:', response.status);
                    return false;
                }
            } catch (error) {
                console.error('[HEALTH] 无法连接到服务器:', error);
                return false;
            }
        }

        // 页面加载时执行健康检查
        window.addEventListener('DOMContentLoaded', async () => {
            const isHealthy = await checkServerHealth();
            if (!isHealthy) {
                console.warn('[WARN] 服务器连接检查失败，切片功能可能无法正常使用');
                showError('警告: 无法连接到服务器，切片功能可能无法正常工作。请检查服务器是否已启动。');
            }
        });

        // 层级选项和提示
        const levelOptions = document.querySelectorAll('input[name="sliceLevel"]');
        const levelHint = document.getElementById('levelHint');
        const levelHints = {
            '1': '只按一级大纲切分，二级及以下标题合并到一级章节',
            '2': '最多切到二级标题，三级及以下合并到二级章节',
            '3': '最多切到三级标题，四级及以下合并到三级章节',
            'all': '按全部大纲层级切分，所有级别标题都成为独立章节'
        };

        // 监听层级选择变化
        levelOptions.forEach(option => {
            option.addEventListener('change', (e) => {
                levelHint.textContent = levelHints[e.target.value] || levelHints['2'];
            });
        });

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
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                handleFile(files[0]);
            }
        });

        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length > 0) {
                handleFile(e.target.files[0]);
            }
        });

        function handleFile(file) {
            if (!file.name.endsWith('.docx')) {
                showError('请上传 .docx 格式的文件');
                return;
            }
            selectedFile = file;
            const fileSizeInfo = `✓ 已选择: <strong>${file.name}</strong><br>大小: ${formatFileSize(file.size)}`;
            fileInfo.innerHTML = fileSizeInfo;
            fileInfo.classList.add('show');
            sliceBtn.disabled = false;
            result.innerHTML = '';
            result.classList.remove('show');

            // 大文件提示
            if (file.size > 50 * 1024 * 1024) { // 大于50MB
                const warningDiv = document.createElement('div');
                warningDiv.style.cssText = 'margin-top: 10px; padding: 10px; background: #fff7e6; border-left: 3px solid #faad14; border-radius: 4px; color: #856404; font-size: 13px; line-height: 1.5;';
                warningDiv.innerHTML = '<strong>⚠️ 大文件提示:</strong><br>文件较大，处理可能需要较长时间。请耐心等待，不要关闭页面。如果上传失败，建议使用桌面版工具。';
                fileInfo.innerHTML += warningDiv.outerHTML;
            }

            // 移除旧的固定下载按钮
            const oldBtn = document.querySelector('.download-btn-fixed');
            if (oldBtn) {
                oldBtn.remove();
            }
        }

        function formatFileSize(bytes) {
            if (bytes < 1024) return bytes + ' B';
            if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(2) + ' KB';
            return (bytes / (1024 * 1024)).toFixed(2) + ' MB';
        }

        function showError(message) {
            // 添加帮助链接和故障排查建议
            const helpText = `
                <div style="margin-top: 10px; font-size: 12px; color: #999;">
                    <strong>故障排查建议:</strong><br>
                    • 检查浏览器控制台 (F12) 获取详细错误信息<br>
                    • 确认服务器正在运行 (访问 <a href="/health" target="_blank" style="color: #667eea; text-decoration: underline;">/health</a>)<br>
                    • 尝试刷新页面重新加载<br>
                    • 如果问题持续，请使用桌面版工具
                </div>
            `;

            result.innerHTML = `<div class="error">${message}${helpText}</div>`;
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

            // 添加切分层级参数
            const selectedLevel = document.querySelector('input[name="sliceLevel"]:checked').value;
            formData.append('max_level', selectedLevel);

            updateProgress(10, '正在上传...');

            // 创建 AbortController 用于超时控制
            const controller = new AbortController();
            // 超时时间根据文件大小动态调整：大文件给予更长时间
            const timeoutMs = Math.max(300000, selectedFile.size * 0.001); // 至少5分钟，大文件给予更多时间
            const timeoutId = setTimeout(() => controller.abort(), timeoutMs);
            console.log(`[DEBUG] 设置超时时间: ${Math.round(timeoutMs / 1000)}秒`);

            try {
                const response = await fetch('/slice', {
                    method: 'POST',
                    body: formData,
                    signal: controller.signal
                });

                clearTimeout(timeoutId);

                const levelText = selectedLevel === 'all' ? '全部' : selectedLevel + '级';
                updateProgress(80, `正在切片（${levelText}）...`);

                if (!response.ok) {
                    let errorMsg = '处理失败';
                    // 413 状态码表示文件过大
                    if (response.status === 413) {
                        errorMsg = '文件大小超过限制，最大支持 2 GB (2048 MB)';
                    } else if (response.status === 400) {
                        // 400 状态码可能包含客户端断开连接错误
                        try {
                            const errorText = await response.text();
                            try {
                                const errorData = JSON.parse(errorText);
                                errorMsg = errorData.error || errorMsg;
                            } catch (e) {
                                errorMsg = errorText || `服务器错误 (${response.status})`;
                            }
                        } catch (e) {
                            errorMsg = `服务器错误 (${response.status})`;
                        }
                    } else {
                        try {
                            const errorText = await response.text();
                            try {
                                const errorData = JSON.parse(errorText);
                                errorMsg = errorData.error || errorMsg;
                            } catch (e) {
                                // 如果不是JSON格式，直接使用文本
                                errorMsg = errorText || `服务器错误 (${response.status})`;
                            }
                        } catch (e) {
                            errorMsg = `服务器错误 (${response.status})`;
                        }
                    }
                    throw new Error(errorMsg);
                }

                updateProgress(100, '处理完成！');

                // 先检查响应的 Content-Type
                const contentType = response.headers.get('Content-Type');
                if (!contentType || !contentType.includes('application/zip')) {
                    const errorText = await response.text();
                    throw new Error(`服务器返回了非ZIP文件: ${errorText.substring(0, 200)}`);
                }

                let blob;
                try {
                    blob = await response.blob();
                    if (blob.size === 0) {
                        throw new Error('服务器返回了空文件');
                    }
                } catch (e) {
                    throw new Error(`无法处理服务器响应: ${e.message}`);
                }
                const url = URL.createObjectURL(blob);

                result.innerHTML = `
                    <div class="success">
                        ✅ 切片完成！<br>
                        共 ${response.headers.get('X-Section-Count') || '多个'} 个章节
                    </div>
                `;

                // 在右下角显示固定下载按钮
                const downloadBtn = document.createElement('a');
                downloadBtn.href = url;
                downloadBtn.download = 'sliced_documents.zip';
                downloadBtn.className = 'download-btn download-btn-fixed';
                downloadBtn.innerHTML = '⬇️ 下载切片结果';
                document.body.appendChild(downloadBtn);
                result.classList.add('show');

            } catch (error) {
                clearTimeout(timeoutId);
                console.error('[ERROR] 请求失败:', error);
                console.error('[ERROR] 错误名称:', error.name);
                console.error('[ERROR] 错误消息:', error.message);
                console.error('[ERROR] 错误堆栈:', error.stack);
                console.error('[DEBUG] 当前页面URL:', window.location.href);

                let errorMsg = error.message;
                let errorDetails = '';

                if (error.name === 'AbortError') {
                    errorMsg = '请求超时，文件可能太大或处理时间过长，请尝试较小的文件或使用桌面版工具';
                } else if (error.name === 'TypeError') {
                    errorMsg = '网络连接失败 - 服务器可能未启动或无法访问';
                    errorDetails = '请检查: 1) 服务器是否正常运行 2) 浏览器控制台是否有详细错误 3) 防火墙是否阻止了连接';
                } else if (error.message && (error.message.includes('Failed to fetch') || error.message.includes('Load failed') || error.message.includes('请求正文流已耗尽') || error.message.includes('请求流') || error.message.includes('stream'))) {
                    errorMsg = '上传中断 - 请求正文流已耗尽';
                    errorDetails = '大文件上传时可能遇到网络不稳定问题。建议: 1) 检查网络连接稳定性 2) 尝试使用桌面版工具处理大文件 3) 将文件分成较小的部分处理';
                } else if (error.message && error.message.includes('NetworkError')) {
                    errorMsg = '网络错误';
                    errorDetails = '请检查网络连接和服务器状态';
                }

                showError(`❌ ${errorMsg}${errorDetails ? '\n\n详情: ' + errorDetails : ''}`);
            }

            sliceBtn.disabled = false;
        });

        function updateProgress(value, text) {
            progressFill.style.width = value + '%';
            progressText.textContent = text;
        }
    </script>
</body>
</html>
    ''')


@app.route('/health')
def health_check():
    """健康检查端点"""
    return jsonify({
        'status': 'ok',
        'timestamp': datetime.now().isoformat(),
        'version': '1.0.0'
    })


@app.route('/info')
def server_info():
    """服务器信息端点"""
    return jsonify({
        'max_upload_size_mb': app.config['MAX_CONTENT_LENGTH'] // (1024 * 1024),
        'supported_formats': ['.docx'],
        'server_time': datetime.now().isoformat()
    })


@app.route('/slice', methods=['POST'])
def slice_file():
    """切片文件"""
    upload_path = None
    try:
        print(f"[DEBUG] ===== /slice 请求开始 =====")
        print(f"[DEBUG] 收到请求，内容长度: {request.content_length}")
        print(f"[DEBUG] 最大允许: {app.config['MAX_CONTENT_LENGTH']}")
        print(f"[DEBUG] 请求方法: {request.method}")
        print(f"[DEBUG] 请求路径: {request.path}")

        print(f"[DEBUG] 收到请求，文件列表: {list(request.files.keys())}")
        if 'file' not in request.files:
            return jsonify({'error': '未上传文件'}), 400

        file = request.files['file']
        print(f"[DEBUG] 文件名: {file.filename}")
        if file.filename == '':
            return jsonify({'error': '文件名为空'}), 400

        if not file.filename.endswith('.docx'):
            return jsonify({'error': '只支持 .docx 格式'}), 400

        print(f"[DEBUG] 开始保存文件到: {UPLOAD_FOLDER}")

        # 获取切分层级参数
        max_level = request.form.get('max_level', '0')
        try:
            if max_level.lower() == 'all':
                max_level = 0  # 0表示全部层级
            else:
                max_level = int(max_level)
                if max_level not in [0, 1, 2, 3]:
                    return jsonify({'error': '无效的切分层级'}), 400
        except ValueError:
            return jsonify({'error': '切分层级参数格式错误'}), 400

        # 保存上传的文件
        upload_path = UPLOAD_FOLDER / file.filename
        print(f"[DEBUG] 保存文件到: {upload_path}")
        file.save(str(upload_path))
        print(f"[DEBUG] 文件已保存，大小: {upload_path.stat().st_size if upload_path.exists() else 0} bytes")

        # 切片
        print(f"[DEBUG] 开始切片，max_level={max_level}")
        slicer = TenderSlicer(upload_path)
        sections = slicer.slice_document(max_level=max_level if max_level > 0 else None)
        print(f"[DEBUG] 切片完成，共 {len(sections)} 个章节")

        # 创建 ZIP 文件
        zip_buffer = io.BytesIO()

        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # 保存各个章节
            for section in sections:
                index = section['index'] + 1
                index_str = str(index).zfill(3)
                safe_title = slicer.sanitize_filename(section['title'])
                filename = f"{index_str}_{safe_title}.md"
                content = ''.join(section['content'])
                zipf.writestr(filename, content.encode('utf-8'))

            # 生成索引
            index_content = "# 标书切片索引\\n\\n"
            index_content += f"原文件: {file.filename}\\n"
            index_content += f"切片时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\\n"
            index_content += f"总章节数: {len(sections)}\\n\\n---\\n\\n## 章节列表\\n\\n"

            for section in sections:
                idx = section['index'] + 1
                level = section['level']
                title = section['title']
                filename = f"{str(idx).zfill(3)}_{slicer.sanitize_filename(title)}.md"
                indent = "  " * (level - 1)
                index_content += f"{indent}- [{idx}. {title}]({filename})\\n"

            zipf.writestr("00_index.md", index_content.encode('utf-8'))

        zip_buffer.seek(0)

        # 删除临时文件
        try:
            upload_path.unlink()
        except Exception:
            pass  # 忽略删除错误

        # 使用 URL 编码处理中文文件名
        from urllib.parse import quote
        safe_filename = quote(f"sliced_{file.filename}.zip", safe='')
        response = app.response_class(
            zip_buffer.getvalue(),
            mimetype='application/zip',
            headers={
                'Content-Disposition': f"attachment; filename*=UTF-8''{safe_filename}",
                'X-Section-Count': str(len(sections))
            }
        )

        return response

    except RequestEntityTooLarge as e:
        # 处理文件过大错误
        print(f"[ERROR] 文件过大: {e}")
        max_size_mb = app.config['MAX_CONTENT_LENGTH'] // (1024 * 1024)
        return jsonify({
            'error': f'文件大小超过限制，最大支持 {max_size_mb} MB'
        }), 413
    except ClientDisconnected as e:
        # 处理客户端断开连接错误（通常是因为文件太大或网络问题）
        print(f"[ERROR] 客户端断开连接: {e}")
        return jsonify({
            'error': '上传中断，可能是网络不稳定或文件太大，请使用桌面版工具'
        }), 400
    except Exception as e:
        # 打印完整的错误堆栈
        import traceback
        print(f"[ERROR] 切片失败: {str(e)}")
        print(f"[ERROR] 堆栈跟踪:")
        traceback.print_exc()

        # 清理临时文件
        if upload_path and upload_path.exists():
            try:
                upload_path.unlink()
            except Exception:
                pass

        # 返回详细的错误信息
        error_msg = str(e)
        if len(error_msg) > 500:
            error_msg = error_msg[:500] + '...'
        return jsonify({'error': error_msg}), 500


if __name__ == '__main__':
    import socket

    # 检查端口是否可用
    def check_port_available(port):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(('0.0.0.0', port))
                return True
            except OSError:
                return False

    print("=" * 50)
    print("标书切片工具 - Web 版本")
    print("=" * 50)
    print("正在启动服务器...")

    # 检查端口 8000 是否被占用
    SERVER_PORT = 8000
    if not check_port_available(SERVER_PORT):
        print("[ERROR] 端口 8000 已被占用！")
        print("[ERROR] 请先停止占用该端口的进程:")
        print(f"[ERROR]   lsof -ti:{SERVER_PORT} | xargs kill -9")
        print("[ERROR] 或者使用以下命令查看占用进程:")
        print(f"[ERROR]   lsof -i:{SERVER_PORT}")
        import sys
        sys.exit(1)
    print("请在浏览器中访问: http://localhost:8000")
    print("按 Ctrl+C 停止服务器")
    print(f"最大上传文件大小: {app.config['MAX_CONTENT_LENGTH'] // (1024 * 1024)} MB (2GB)")
    print("=" * 50)

    # 使用 threaded=True 支持并发请求，后台运行时关闭 debug 模式
    app.run(host='0.0.0.0', port=8000, debug=False, threaded=True)
