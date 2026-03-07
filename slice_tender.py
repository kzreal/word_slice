#!/usr/bin/env python3
"""
标书文档切片脚本
将 Word (.docx) 文档按目录大纲结构切片为多个编号的 Markdown 文件
"""

import os
import re
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn


class TenderSlicer:
    """标书切片器"""

    def __init__(self, docx_path: str, output_dir: str = "sliced"):
        self.docx_path = Path(docx_path)
        self.output_dir = Path(output_dir)
        self.doc = None
        self.sections = []  # 存储切片结果

        # 创建输出目录
        self.output_dir.mkdir(exist_ok=True)

    def load_document(self):
        """加载 Word 文档"""
        if not self.docx_path.exists():
            raise FileNotFoundError(f"文件不存在: {self.docx_path}")
        self.doc = Document(str(self.docx_path))

    def get_heading_level(self, paragraph):
        """
        获取段落的标题级别
        返回: 0=正文, 1=一级标题, 2=二级标题, ...
        """
        style_name = paragraph.style.name
        if 'Heading' in style_name:
            # 提取标题级别，如 "Heading 1" -> 1
            match = re.search(r'Heading\s*(\d+)', style_name, re.IGNORECASE)
            if match:
                return int(match.group(1))
        # 检查大纲级别（有些文档使用大纲级别而不是标题样式）
        if hasattr(paragraph, '_element'):
            p = paragraph._element
            if p.pPr and p.pPr.outlineLvl is not None:
                # outlineLvl 是从0开始的，所以加1
                return int(p.pPr.outlineLvl.val) + 1
        return 0

    def get_paragraph_text(self, paragraph):
        """获取段落文本（包括表格内容）"""
        text = paragraph.text.strip()
        if text:
            return text
        return None

    def slice_document(self):
        """
        切片文档
        按照标题级别进行切片
        """
        if not self.doc:
            self.load_document()

        current_section = {
            'level': 0,
            'title': '封面',
            'content': [],
            'index': 0
        }
        section_index = 0

        for paragraph in self.doc.paragraphs:
            level = self.get_heading_level(paragraph)
            text = self.get_paragraph_text(paragraph)

            # 跳过空段落
            if not text:
                continue

            # 检测目录页（跳过）
            if any(keyword in text for keyword in ['目录', '目  录', 'CONTENTS', 'TABLE OF CONTENTS']):
                continue

            if level > 0:
                # 保存当前章节
                if current_section['content']:
                    self.sections.append(current_section)
                    section_index += 1

                # 开始新章节
                current_section = {
                    'level': level,
                    'title': text,
                    'content': [],
                    'index': section_index
                }
                current_section['content'].append(f"{'#' * level} {text}\n")
            else:
                current_section['content'].append(f"{text}\n")

        # 处理表格
        for table in self.doc.tables:
            table_md = self.table_to_markdown(table)
            if table_md and current_section['content']:
                current_section['content'].append(table_md)

        # 保存最后一个章节
        if current_section['content']:
            self.sections.append(current_section)

    def table_to_markdown(self, table):
        """将表格转换为 Markdown 格式"""
        if not table.rows:
            return None

        markdown = []
        for i, row in enumerate(table.rows):
            cells = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
            markdown.append('| ' + ' | '.join(cells) + ' |')

            # 添加分隔线（第一行之后）
            if i == 0:
                separator = '|' + '|'.join(['---'] * len(cells)) + '|'
                markdown.append(separator)

        return '\n'.join(markdown) + '\n\n'

    def sanitize_filename(self, filename):
        """清理文件名，移除非法字符"""
        # 移除或替换非法字符
        illegal_chars = r'[<>:"/\\|?*]'
        filename = re.sub(illegal_chars, '', filename)
        # 限制长度
        if len(filename) > 100:
            filename = filename[:100]
        return filename.strip()

    def save_sections(self):
        """保存切片后的文件"""
        print(f"正在保存 {len(self.sections)} 个章节...")

        for section in self.sections:
            # 生成编号
            index = section['index'] + 1
            index_str = str(index).zfill(3)  # 补零到3位

            # 清理标题作为文件名
            safe_title = self.sanitize_filename(section['title'])
            filename = f"{index_str}_{safe_title}.md"
            filepath = self.output_dir / filename

            # 写入内容
            content = ''.join(section['content'])
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(content)

            print(f"  已保存: {filename}")

        print(f"\n完成！所有文件已保存到: {self.output_dir}")

    def generate_index(self):
        """生成索引文件"""
        index_file = self.output_dir / "00_index.md"
        with open(index_file, 'w', encoding='utf-8') as f:
            f.write("# 标书切片索引\n\n")
            f.write(f"原文件: {self.docx_path.name}\n")
            f.write(f"切片时间: {self._get_timestamp()}\n")
            f.write(f"总章节数: {len(self.sections)}\n\n")
            f.write("---\n\n")
            f.write("## 章节列表\n\n")

            for section in self.sections:
                index = section['index'] + 1
                level = section['level']
                title = section['title']
                filename = f"{str(index).zfill(3)}_{self.sanitize_filename(title)}.md"

                # 根据级别添加缩进
                indent = "  " * (level - 1)
                f.write(f"{indent}- [{index}. {title}]({filename})\n")

        print(f"索引文件已生成: {index_file}")

    @staticmethod
    def _get_timestamp():
        """获取当前时间戳"""
        from datetime import datetime
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def process(self):
        """执行完整的切片流程"""
        print(f"开始处理: {self.docx_path}")
        print("-" * 50)
        self.load_document()
        self.slice_document()
        self.save_sections()
        self.generate_index()
        print("-" * 50)
        print("处理完成！")


def main():
    """主函数"""
    import sys

    if len(sys.argv) < 2:
        print("使用方法: python slice_tender.py <docx文件路径> [输出目录]")
        print("\n示例:")
        print("  python slice_tender.py 投标文件.docx")
        print("  python slice_tender.py 投标文件.docx sliced_output")
        sys.exit(1)

    docx_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else "sliced"

    try:
        slicer = TenderSlicer(docx_path, output_dir)
        slicer.process()
    except Exception as e:
        print(f"错误: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
