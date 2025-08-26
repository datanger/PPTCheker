# 🤖 LLM智能审查使用指南

## 🎯 功能概述

基于大模型的智能PPT审查系统，相比传统规则检查具有以下优势：

### ✨ **智能特性**
- **语义理解**：理解PPT内容的上下文和含义
- **灵活判断**：基于内容智能判断是否需要修复
- **个性化建议**：提供具体的改进方案和替代表达
- **多维度分析**：格式、逻辑、术语、流畅性全方位审查

### 🔄 **混合架构**
- **规则引擎**：处理明确的格式规范问题
- **LLM引擎**：处理语义化和上下文相关的问题
- **智能降级**：LLM不可用时自动降级为纯规则模式

## 🚀 快速开始

### 1. 环境配置
```bash
# 设置 DeepSeek API 密钥（推荐）
export DEEPSEEK_API_KEY="your_deepseek_api_key"

# 或者使用通用环境变量
export LLM_API_KEY="your_api_key"
export LLM_MODEL="deepseek-chat"
export LLM_ENDPOINT="https://api.deepseek.com/v1/chat/completions"

# 或使用配置文件
cp configs/config_llm.yaml configs/config.yaml
```

### 2. 运行智能审查
```bash
# 在app目录下运行
cd app
python -m pptlint.cli --input "../智能体及扣子介绍.pptx" --config "../configs/config.yaml" --output-ppt "../out/智能审查版.pptx"

# 或使用WebUI（推荐）
streamlit run webui/app.py
```

## 📋 审查维度详解

### 1. **格式规范审查** (`LLM_FormatRule`)
- **字体检查**：智能判断字体使用是否合适
- **字号检查**：基于内容重要性推荐字号
- **颜色检查**：分析颜色搭配的协调性
- **布局检查**：评估页面布局的合理性

### 2. **内容逻辑审查** (`LLM_ContentRule`)
- **逻辑连贯性**：检查各页面间的逻辑过渡
- **内容完整性**：识别可能遗漏的重要信息
- **结构合理性**：评估PPT整体结构设计
- **重点突出**：分析关键信息的表达效果

### 3. **缩略语审查** (`LLM_AcronymRule`)
- **智能识别**：基于上下文判断缩略语是否需要解释
- **解释建议**：提供具体的解释方式和位置建议
- **一致性检查**：确保相同概念使用统一术语

### 4. **表达流畅性审查** (`LLM_FluencyRule`)
- **语言表达**：检查文字表达的清晰度和准确性
- **专业术语**：确保术语使用的专业性和一致性
- **表达风格**：保持整体表达风格的统一

## ⚙️ 配置说明

### 基础配置
```yaml
# LLM开关
llm_enabled: true

# 模型设置
llm_model: "deepseek-chat"
llm_temperature: 0.2        # 创造性，0.0-1.0
llm_max_tokens: 1024       # 最大输出长度

# 审查维度开关
review_format: true         # 格式审查
review_logic: true          # 逻辑审查
review_acronyms: true       # 缩略语审查
review_fluency: true        # 流畅性审查
```

### 环境变量配置
```bash
# 必需配置
export LLM_API_KEY="your_api_key"

# 可选配置
export LLM_MODEL="deepseek-chat"
export LLM_ENDPOINT="https://api.deepseek.com/v1/chat/completions"
```

## 📊 输出示例

### 智能审查报告
```markdown
## 审查报告

### 📊 问题统计
- **规则检查问题**: 143 个
- **LLM智能审查问题**: 25 个
- **总计**: 168 个

### 🔍 规则检查问题
- **FontFamilyRule** | 严重性: warning | 页: 0 | 对象: 7
  - 描述: 日文字体非 Meiyou UI: 未指定
  - 建议: 替换为 Meiyou UI

### 🤖 LLM智能审查问题
- **LLM_ContentRule** | 严重性: info | 页: 2 | 对象: page
  - 描述: 第2页与第3页之间缺乏逻辑过渡，建议添加过渡语句
  - 建议: 在页面末尾添加"接下来我们将详细介绍..."等过渡语

- **LLM_AcronymRule** | 严重性: warning | 页: 1 | 对象: page
  - 描述: "API"首次出现缺乏解释，但根据上下文判断需要解释
  - 建议: 在"API"后添加"(Application Programming Interface)"解释
```

## 🔧 故障排除

### 常见问题

#### 1. **LLM服务不可用**
```
ℹ️ LLM未配置，使用纯规则审查模式
```
**解决方案**：检查API密钥和网络连接

#### 2. **LLM审查失败**
```
⚠️ LLM审查失败，降级为纯规则模式
```
**解决方案**：检查模型名称和API端点配置

#### 3. **响应解析错误**
```
LLM格式审查失败: Expecting value: line 1 column 1
```
**解决方案**：检查LLM返回的JSON格式是否正确

### 调试模式
```bash
# 启用详细日志
export PYTHONPATH=.
python -m pptlint.cli --input "test.pptx" --config "config.yaml" --llm on
```

## 🎯 最佳实践

### 1. **提示词优化**
- 使用清晰、具体的审查标准
- 提供足够的上下文信息
- 明确输出格式要求

### 2. **配置调优**
- 根据文档类型调整审查维度
- 平衡审查深度和性能
- 定期更新审查标准

### 3. **结果验证**
- 人工复核LLM建议
- 结合业务需求判断
- 持续优化审查效果

## 🚧 未来规划

### 短期目标
- [ ] 支持更多LLM模型
- [ ] 优化提示词工程
- [ ] 增加审查维度

### 长期目标
- [ ] 支持多语言审查
- [ ] 实现智能自动修复
- [ ] 集成行业专业知识库

## 📞 技术支持

如有问题或建议，请：
1. 查看日志输出
2. 检查配置参数
3. 提交Issue反馈
4. 联系开发团队
