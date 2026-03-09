# 标书切片工具 - 切片逻辑说明

## 整体架构

采用"先转换，再切片"的两阶段处理模式：

```
Word 文档
    ↓
阶段一：转换 (convert_to_markdown)
    ↓
完整的 Markdown 内容
    ↓
阶段二：切片 (_slice_markdown_by_level)
    ↓
切片后的章节列表
```

## 核心方法

### 1. convert_to_markdown()
将整个 Word 文档转换成 Markdown 格式。

**处理流程**：
1. 收集所有段落和表格，按它们在文档中的位置排序
2. 遍历元素，按原文顺序转换成 Markdown
3. 保留原始标题结构（`#`、`##`、`###`）
4. 跳过空段落和目录页
5. 返回完整的 Markdown 字符串

**位置排序**：通过 `element._element.getparent().index(element._element)` 获取元素在文档 body 中的位置索引，确保表格被正确归位，不会全部堆积到最后一个章节。

### 2. _parse_markdown_structure()
解析 Markdown 内容，提取所有标题信息。

**返回**：
- 标题列表，每个元素包含：`{level, title, start_line}`
- 所有行的列表

### 3. _slice_markdown_by_level(md_content, slice_level)
根据切片级别对 Markdown 进行切片。

**切片级别说明**：

| 级别 | 说明 | 行为 |
|------|------|------|
| 0 | 零级模式 | 不切片，整个文档为一个 Markdown 文件 |
| None | 全部模式 | 按所有标题层级切片 |
| 1 | 一级 | 只按一级标题（#）切片 |
| 2 | 二级 | 只按二级标题（##）切片 |
| 3 | 三级 | 只按三级标题（###）切片 |

**切片逻辑**：
1. 解析 Markdown 中的所有标题
2. 根据标题级别确定章节边界
3. 提取每个章节的内容范围
4. 返回章节列表

**章节边界确定**：
- 如果标题级别 <= max_level，则创建新章节
- 否则，该标题及其内容归属于上一个章节

### 4. slice_document()
主入口方法，协调整个切片流程。

**步骤**：
1. 调用 `convert_to_markdown()` 转换成 Markdown
2. 调用 `_slice_markdown_by_level()` 进行切片
3. 将结果存储到 `self.sections`

## 使用示例

```bash
# 零级模式 - Word 转成一个 md，不切片
python slice_tender.py 投标文件.docx sliced_output 0

# 一级切片 - 按一级标题切片
python slice_tender.py 投标文件.docx sliced_output 1

# 二级切片 - 按二级标题切片
python slice_tender.py 投标文件.docx sliced_output 2

# 全部切片 - 按所有标题层级切片
python slice_tender.py 投标文件.docx sliced_output
```

## 优势

1. **职责分离**：转换和切片逻辑独立，易于维护和测试
2. **易于调试**：可以查看完整的 Markdown 中间结果
3. **保持顺序**：通过位置排序确保表格和内容按原文顺序排布
4. **灵活切片**：支持多级别切片，零级模式满足"Word 转 MD"需求

## 注意事项

1. 零级模式输出的 Markdown 文件名使用原文件名（去除 .docx 扩展名）
2. 目录页会被自动跳过（中文"目录"、"目  录"，英文"CONTENTS"、"TABLE OF CONTENTS"）
3. 表格会被正确转换并保持在原文位置
4. 所有章节文件按 `001_标题.md` 格式命名，便于排序
