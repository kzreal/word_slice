# 图片格式改造方案 v2.0

## 需求分析

### 主要需求
1. **格式调整**：将图片格式从 `<!-- 89 [图片] 内容: 描述 -->` 改为 `<!-- 89 -->[图片: 描述]`
2. **表格图片合并**：将表格同一行的多张图片合并为一行输出

### 示例

**当前格式：**
```markdown
|---|---|
|---|---|
<!-- 89 [图片] 内容: 身份证背面 -->
<!-- 90 [图片] 内容: 身份证正面，姓名：zxm -->
```

**期望格式：**
```markdown
|---|---|
|---|---|
<!-- 89 -->[图片: 身份证背面] [图片: 身份证正面，姓名：zxm]
<!-- 90 -->普通文本内容
```

## 可行性分析

### ✅ 技术可行性

#### 1. 图片格式调整
- 原格式：`<!-- {行号} [图片] 内容: {描述} -->`
- 新格式：`<!-- {行号} -->[图片: {描述}]`
- 实现简单，只需修改字符串拼接方式

#### 2. 表格图片合并
- **挑战**：当前表格每张图片都独立占用一个行号
- **解决方案**：
  1. 先收集表格中所有图片及其位置
  2. 按行分组图片
  3. 同一行内的图片合并输出，只占用一个行号
  4. 该行非图片内容继续使用下一个行号

### ✅ 实现方案

#### 1. 图片格式统一调整
在所有四个图片输出位置修改：
```python
# 修改前
full_section['content'].append(f"<!-- {line_no} [图片] 内容: {description} -->\n")

# 修改后
full_section['content'].append(f"<!-- {line_no} -->[图片: {description}]\n")
```

#### 2. 表格图片合并处理

**新增数据结构：**
```python
# 用于存储表格行的图片信息
class TableRowImages:
    def __init__(self, line_no):
        self.line_no = line_no  # 行号
        self.images = []       # 该行的图片列表
```

**处理流程：**
1. 遍历表格时，记录每行的图片
2. 处理完一行后，如果有图片：
   - 合并该行所有图片：`[图片: 描述1] [图片: 描述2]`
   - 只使用一个行号输出合并后的图片
   - 跳过后续图片占用的行号

**代码实现：**
```python
# 表格处理逻辑
current_row_images = TableRowImages(line_no)
for row in table.rows:
    # 收集当前行的所有图片
    for cell in row.cells:
        # 提取图片并添加到 current_row_images

    if current_row_images.images:
        # 合并输出
        descriptions = []
        for img in current_row_images.images:
            desc = processed_images.get(img['id'])
            if desc:
                descriptions.append(f"[图片: {desc}]")

        content = f"<!-- {current_row_images.line_no} -->{' '.join(descriptions)}\n"
        full_section['content'].append(content)

        # 跳过图片占用的行号
        line_no += len(descriptions)
    else:
        # 普通表格行，正常处理
        line_no += 1
```

### ✅ 实现步骤

1. **修改图片格式**（所有四个位置）
   - 将 `<!-- {行号} [图片] 内容: 描述 -->` 改为 `<!-- {行号} -->[图片: 描述]`

2. **重构表格处理逻辑**
   - 修改 `table_to_markdown` 方法，使其支持图片合并
   - 或者创建新的 `table_to_markdown_with_images` 方法
   - 在 `slice_document` 中使用新的表格处理逻辑

3. **保持行号连续性**
   - 合并的图片只占用一个行号
   - 后续内容继续使用递增的行号

4. **兼容性处理**
   - 确保无图片的表格行正常处理
   - 确保跨行的图片不被错误合并

### ✅ 优势

1. **更紧凑的布局**
   - 同一行的图片合并显示
   - 减少不必要的空行

2. **更好的可读性**
   - 行号和图片内容分离
   - 图片描述更直观

3. **保持结构清晰**
   - HTML 注释只包含行号
   - 图片内容在 `[图片: ]` 中明确标记

### ⚠️ 注意事项

1. **复杂表格处理**
   - 合并单元格中的图片需要特殊处理
   - 确保图片的顺序保持正确

2. **行号计算**
   - 合并图片后要正确计算后续内容的行号
   - 避免行号跳跃或重复

3. **测试用例**
   - 需要测试表格中有/无图片的情况
   - 测试跨行图片的情况
   - 测试混合内容（文本+图片）的情况

### 🔧 具体实现建议

1. **创建辅助方法**
   ```python
   def format_images_in_row(images, line_no):
       """格式化同一行的多张图片"""
       if not images:
           return None, 0

       descriptions = []
       for img in images:
           desc = processed_images.get(img['id'])
           if desc:
               descriptions.append(f"[图片: {desc}]")

       content = f"<!-- {line_no} -->{' '.join(descriptions)}\n"
       return content, len(descriptions)
   ```

2. **修改表格处理循环**
   - 在表格处理时收集行的图片
   - 调用 `format_images_in_row` 合并输出
   - 根据返回的图片数量调整行号

## 结论

这个方案完全可行，实现相对复杂但可以做到。主要改动在表格处理逻辑，图片格式调整比较简单。需要特别注意行号计算的准确性。