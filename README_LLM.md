# LLM 图片识别功能使用说明

## 功能概述

标书切片工具现在支持通过 LLM API 识别文档中的图片内容。识别结果会以 `<!-- [图片] LLM识别结果: [描述] -->` 的形式添加到 Markdown 文件中。

## 安装依赖

```bash
pip install -r requirements.txt
```

## 配置环境变量

### 1. 复制环境变量模板
```bash
cp .env.example .env
```

### 2. 编辑 .env 文件
```bash
# 必填项
LLM_API_ENDPOINT=https://your-api-endpoint.com/v1/analyze-image
LLM_API_KEY=your-api-key-here

# 可选配置
LLM_MODEL=vision-model        # 默认模型
LLM_TIMEOUT=30                # 超时时间（秒）
LLM_MAX_RETRIES=3            # 最大重试次数
```

## 使用方法

### Web 界面使用

1. 启动服务器
```bash
python3 tender_slicer_web.py
```

2. 浏览器访问 http://localhost:8000

3. 上传 Word 文档

4. 切片级别选择：
   - 零级：整个文档转成一个 Markdown 文件，包含所有图片识别
   - 一级/二级/三级：按级别切片，每个切片包含该章节内的图片识别
   - 全部：按所有标题切片，包含所有图片识别

### 命令行测试

```bash
# 测试 LLM 配置
python3 test_llm_integration.py
```

## 工作原理

1. **图片检测**：扫描文档中的段落和表格，检测嵌入的图片
2. **数据提取**：将图片数据提取为二进制格式
3. **LLM 调用**：将图片数据发送到配置的 LLM API
4. **结果处理**：将 LLM 的识别结果替换原来的 `<!-- [图片] -->` 占位符
5. **降级处理**：API 调用失败时使用占位符，保证文档完整性

## 错误处理

### 配置错误
- 缺少环境变量时会显示警告并使用占位符
- 不会中断文档处理

### 网络错误
- 自动重试最多 3 次
- 重试失败后使用占位符

### API 错误
- 4xx 错误：认证失败、请求参数错误等
- 5xx 错误：服务器错误
- 所有错误都会记录日志并使用占位符

## 性能优化

1. **批量处理**：所有图片集中处理，减少 API 调用开销
2. **图片大小限制**：超过 10MB 的图片自动使用占位符
3. **连接复用**：使用 HTTP session 复用连接
4. **超时控制**：防止长时间等待

## 日志信息

处理过程中会记录以下日志：
- `INFO`: LLM 服务启用状态
- `INFO`: 图片处理进度
- `WARNING`: LLM 不可用或失败时的降级处理
- `ERROR`: 详细错误信息

## 故障排除

### 1. LLM API 不可用
```bash
# 检查环境变量
echo $LLM_API_ENDPOINT
echo $LLM_API_KEY
```

### 2. 运行测试脚本
```bash
python3 test_llm_integration.py
```

### 3. 查看日志
```bash
tail -f /tmp/tender_slicer_web.log
```

## 注意事项

1. **API Key 安全**：不要将 API Key 提交到代码仓库
2. **网络要求**：LLM API 必须可访问
3. **图片格式**：支持 PNG、JPEG、GIF、BMP 等常见格式
4. **不保存图片**：只进行识别，不保存原始图片文件

## 支持的 LLM 服务

只要是支持图片输入的 LLM API 都可以使用，需要正确配置请求格式。示例格式：
```json
{
  "model": "vision-model",
  "messages": [
    {
      "role": "user",
      "content": [
        {
          "type": "text",
          "text": "请描述这张图片"
        },
        {
          "type": "image_url",
          "image_url": {
            "url": "data:image/png;base64,..."
          }
        }
      ]
    }
  ]
}
```

## 更新日志

### v1.1.0 (2026-03-09)
- 修复了图片处理中的 bug：
  - 修复了 `line_no` 变量未定义的错误
  - 修复了表格图片处理未使用 LLM 识别结果的问题
  - 统一了图片占位符的格式
  - 优化了错误日志输出
- 添加了测试脚本 `test_llm_fixed.py`
- 改进了错误处理的降级策略

### v1.0.0 (2026-03-09)
- 初始版本发布
- 实现 LLM 图片识别功能
- 支持环境变量配置
- 添加批处理优化
- 完善的错误处理机制