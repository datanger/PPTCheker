# PPT审查工具CLI使用指南

## 🚀 快速开始

### 基础审查模式
```bash
# 生成报告和标记PPT
python -m app.pptlint.cli \
  --parsing parsing_result.json \
  --config configs/config.yaml \
  --mode review \
  --report report.md \
  --output-ppt output.pptx
```

### 仅生成报告
```bash
python -m app.pptlint.cli \
  --parsing parsing_result.json \
  --config configs/config.yaml \
  --mode review \
  --report report.md
```

### 仅生成标记PPT
```bash
python -m app.pptlint.cli \
  --parsing parsing_result.json \
  --config configs/config.yaml \
  --mode review \
  --output-ppt output.pptx
```

## 📋 命令行参数

### 必需参数
- `--parsing`: 解析结果文件路径（parsing_result.json）
- `--config`: 配置文件路径（YAML格式）

### 可选参数
- `--mode`: 运行模式
  - `review` (默认): 审查模式
  - `edit`: 编辑模式
- `--llm`: LLM控制
  - `on`: 启用LLM（默认）
  - `off`: 禁用LLM
- `--report`: 输出报告路径（.md格式）
- `--output-ppt`: 输出PPT路径（.pptx格式）

### 编辑模式专用参数
- `--original-pptx`: 原始PPTX文件路径
- `--edit-req`: 编辑要求提示语

### 高级配置参数（覆盖配置文件设置）
- `--font-size`: 最小字号阈值
- `--color-threshold`: 颜色数量阈值
- **注意**: 缩略语识别由LLM智能进行，无需手动配置

## ⚙️ 配置文件说明

配置文件 `configs/config.yaml` 包含以下设置：

### 字体配置
```yaml
jp_font_name: "Meiryo UI"  # 日文字体统一
min_font_size_pt: 12        # 最小字号（磅）
```

### 缩略语配置
```yaml
# 缩略语识别完全由LLM大模型进行，无需手动设置长度范围
# LLM会智能识别需要解释的专业术语缩略语
```

### 颜色配置
```yaml
color_count_threshold: 5    # 颜色数量阈值
```

### LLM配置
```yaml
llm_enabled: true           # 是否启用LLM审查
llm_model: "deepseek-chat"  # LLM模型
llm_temperature: 0.2        # 温度参数
llm_max_tokens: 1024        # 最大token数
```

### 审查维度开关
```yaml
review_format: true         # 格式规范审查
review_logic: true          # 内容逻辑审查
review_acronyms: true       # 缩略语审查
review_fluency: true        # 表达流畅性审查
```

## 🔧 使用示例

### 1. 基础审查（生成报告和标记PPT）
```bash
python -m app.pptlint.cli \
  --parsing parsing_result.json \
  --config configs/config.yaml \
  --report my_report.md \
  --output-ppt marked_presentation.pptx
```

### 2. 自定义配置参数
```bash
python -m app.pptlint.cli \
  --parsing parsing_result.json \
  --config configs/config.yaml \
  --font-size 14 \
  --color-threshold 3 \
  --report custom_report.md \
  --output-ppt custom_output.pptx
```

### 3. 禁用LLM审查
```bash
python -m app.pptlint.cli \
  --parsing parsing_result.json \
  --config configs/config.yaml \
  --llm off \
  --report no_llm_report.md \
  --output-ppt no_llm_output.pptx
```

### 4. 编辑模式
```bash
python -m app.pptlint.cli \
  --parsing parsing_result.json \
  --config configs/config.yaml \
  --mode edit \
  --original-pptx original.pptx \
  --output-ppt improved.pptx \
  --edit-req "请优化PPT的字体大小和颜色搭配，使其更加美观易读" \
  --report edit_report.md
```

### 5. 仅生成报告
```bash
python -m app.pptlint.cli \
  --parsing parsing_result.json \
  --config configs/config.yaml \
  --report analysis_report.md
```

### 6. 仅生成标记PPT
```bash
python -m app.pptlint.cli \
  --parsing parsing_result.json \
  --config configs/config.yaml \
  --output-ppt marked_presentation.pptx
```

## 🧪 测试CLI功能

运行测试脚本验证CLI功能：
```bash
python test_cli.py
```

测试脚本将验证：
- 基础审查模式
- 仅生成报告模式
- 仅生成标记PPT模式
- 禁用LLM审查模式
- 自定义配置参数模式
- 编辑模式
- 帮助信息显示

## 📊 输出文件说明

### 报告文件 (.md)
- 问题汇总
- 详细问题描述
- 改进建议
- 统计信息

### 标记PPT (.pptx)
- 问题位置标记
- 颜色编码
- 问题说明注释

### 编辑PPT (.pptx)
- 自动修复后的PPT
- 保持原始内容
- 应用改进建议

## 🔍 故障排除

### 常见问题
1. **配置文件不存在**: 确保 `configs/config.yaml` 存在
2. **解析文件不存在**: 确保 `parsing_result.json` 存在
3. **权限问题**: 确保有写入输出目录的权限
4. **LLM连接失败**: 检查网络连接和API配置

### 调试模式
使用 `--help` 查看所有可用参数：
```bash
python -m app.pptlint.cli --help
```

## 📝 注意事项

1. **文件路径**: 使用绝对路径或相对于当前目录的路径
2. **配置文件**: 确保YAML格式正确，避免语法错误
3. **输出目录**: 确保输出目录存在或有权限创建
4. **LLM配置**: 编辑模式需要LLM支持，确保配置正确
5. **文件格式**: 输入必须是JSON，输出支持MD和PPTX

## 🤝 贡献

如有问题或建议，请提交Issue或Pull Request。
