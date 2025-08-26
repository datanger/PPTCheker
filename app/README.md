# PPT 格式审查工具

本项目依据《docs/需求文档_PPT格式审查工具.md》与《docs/架构设计_PPT格式审查工具.md》实现。

## 🚀 快速开始

### 1. 安装依赖
```bash
# 在app目录下安装Python依赖
cd app
pip install -r requirements.txt
```

### 2. CLI 运行
```bash
# 方式1：在app目录下运行（推荐）
cd app
python -m pptlint.cli --input "../智能体及扣子介绍.pptx" --config "../configs/config.yaml" --output-ppt "../out/标记版.pptx"

# 方式2：在项目根目录下运行
python -m app.pptlint.cli --input "智能体及扣子介绍.pptx" --config "configs/config.yaml" --output-ppt "out/标记版.pptx"

# 方式3：生成报告文件
python -m app.pptlint.cli --input "智能体及扣子介绍.pptx" --config "configs/config.yaml" --report "out/report.md" --output-ppt "out/标记版.pptx"
```

### GUI 运行
```bash
# 在app目录下运行
cd app
python -m pptlint.gui
```

### WebUI 运行（推荐，更美观）
```bash
# 在app目录下运行
cd app
streamlit run webui/app.py
```

## 📋 功能概述

### ✅ 已实现功能

#### 1. **基础格式检查**
- **字体/字号检查**：检测最小字号阈值、日文字体统一性
- **英文缩略语解释**：识别2-8字符大写缩略语，检测首次出现是否包含解释
- **颜色数量控制**：单页颜色数量阈值检查
- **主题色调一致性**：预留接口，待完善

#### 2. **智能检测特性**
- **多语言支持**：自动检测日文内容，支持日文字体规范检查
- **智能降级**：LLM服务不可用时自动降级为纯规则检查
- **批量处理**：支持单文件或目录批量处理

#### 3. **输出功能**
- **问题报告**：Markdown格式详细报告
- **PPT标记**：带问题标记的PPT输出（红色斜体下划线标记）
- **问题汇总**：首页全局问题统计

### 🚀 **新增：LLM智能审查功能**

#### 1. **智能格式审查** (`LLM_FormatRule`)
- **智能字体判断**：基于内容上下文判断字体使用是否合适
- **字号推荐**：根据内容重要性智能推荐字号大小
- **颜色协调性**：分析颜色搭配的视觉协调性
- **布局合理性**：评估页面布局设计的合理性

#### 2. **内容逻辑审查** (`LLM_ContentRule`)
- **逻辑连贯性**：检查各页面之间的逻辑过渡是否自然
- **内容完整性**：识别可能遗漏的重要信息点
- **结构合理性**：评估PPT整体结构设计的合理性
- **重点突出**：分析关键信息的表达效果

#### 3. **智能缩略语审查** (`LLM_AcronymRule`)
- **上下文理解**：基于内容语义判断缩略语是否需要解释
- **智能建议**：提供具体的解释方式和位置建议
- **术语一致性**：确保相同概念使用统一术语表达

#### 4. **表达流畅性审查** (`LLM_FluencyRule`)
- **语言表达**：检查文字表达的清晰度和准确性
- **专业术语**：确保术语使用的专业性和一致性
- **表达风格**：保持整体表达风格的统一性

### 🔄 部分实现功能

#### 1. **自动修复**（基础框架已搭建）
- 字体自动修复：配置开关已支持，具体逻辑待完善
- 字号自动修复：配置开关已支持，具体逻辑待完善
- 颜色自动修复：配置开关已支持，具体逻辑待完善
- 缩略语自动修复：当前不支持，需补充实现

#### 2. **LLM增强功能**（接口已预留）
- 日文流畅性建议：接口已预留，具体实现待完善
- 逻辑连贯性建议：接口已预留，具体实现待完善
- 术语统一建议：接口已预留，具体实现待完善

### 📁 目录结构
```
app/
├── pptlint/          # 核心工具源码
│   ├── __init__.py   # 包初始化
│   ├── model.py      # 数据模型定义
│   ├── parser.py     # PPTX解析器
│   ├── rules.py      # 规则引擎
│   ├── workflow.py   # 工作流编排
│   ├── config.py     # 配置管理
│   ├── cli.py        # 命令行接口
│   ├── gui.py        # 桌面GUI
│   ├── llm.py        # LLM集成接口
│   ├── reporter.py   # 报告生成
│   ├── annotator.py  # PPT标记输出
│   └── user_req.py   # 用户需求解析
├── webui/            # Streamlit Web界面
├── configs/          # 配置文件目录
├── dicts/            # 术语词库目录
└── requirements.txt  # Python依赖
```

## ⚙️ 配置说明

### 基础配置 (configs/config.yaml)
```yaml
# 字体设置
jp_font_name: "Meiyou UI"    # 日文字体标准
min_font_size_pt: 12         # 最小字号阈值

# 缩略语设置
acronym_min_len: 2           # 缩略语最小长度
acronym_max_len: 8           # 缩略语最大长度

# 颜色设置
color_count_threshold: 5     # 单页颜色数量阈值

# 自动修复开关
autofix_font: false          # 字体自动修复
autofix_size: false          # 字号自动修复
autofix_color: false         # 颜色自动修复
```

### LLM配置
支持环境变量配置：
```bash
# 推荐：使用 DeepSeek 专用环境变量
export DEEPSEEK_API_KEY="your_deepseek_api_key"

# 或者：使用通用环境变量
export LLM_API_KEY="your_api_key"
export LLM_MODEL="deepseek-chat"
export LLM_ENDPOINT="https://api.deepseek.com/v1/chat/completions"
```

#### LLM智能审查配置
```yaml
# configs/config_llm.yaml
llm_enabled: true
llm_model: "deepseek-chat"
llm_temperature: 0.2
llm_max_tokens: 1024

# 审查维度开关
review_format: true      # 格式规范审查
review_logic: true       # 内容逻辑审查
review_acronyms: true    # 缩略语审查
review_fluency: true     # 表达流畅性审查
```

#### 支持的LLM模型
- **DeepSeek**: `deepseek-chat` (默认推荐)
- **OpenAI**: `gpt-3.5-turbo`, `gpt-4`
- **其他**: 支持OpenAI兼容的API端点

## 🔧 使用方式

### 1. CLI模式（适合批处理）
```bash
# 基础审查（在项目根目录下运行）
python -m app.pptlint.cli --input "input.pptx" --config "configs/config.yaml" --output-ppt "out/标记版.pptx"

# 带报告输出的审查
python -m app.pptlint.cli --input "input.pptx" --config "configs/config.yaml" --report "out/report.md" --output-ppt "out/标记版.pptx"

# 批量处理目录
python -m app.pptlint.cli --input "./ppts/" --config "configs/config.yaml" --report "out/batch_report.md"

# 编辑模式（当前与审查模式相同，自动修复功能待完善）
python -m app.pptlint.cli --input "input.pptx" --config "configs/config.yaml" --mode edit --output-ppt "out/修复版.pptx"

# 在app目录下运行（推荐，路径更简单）
cd app
python -m pptlint.cli --input "../input.pptx" --config "../configs/config.yaml" --output-ppt "../out/标记版.pptx"

#### LLM智能审查模式
```bash
# 启用LLM智能审查（推荐：使用 DeepSeek 专用环境变量）
export DEEPSEEK_API_KEY="your_deepseek_api_key"
python -m app.pptlint.cli --input "input.pptx" --config "configs/config_llm.yaml" --llm on --output-ppt "out/智能审查版.pptx"

# 或者使用通用环境变量
export LLM_API_KEY="your_api_key"
python -m app.pptlint.cli --input "input.pptx" --config "configs/config_llm.yaml" --llm on --output-ppt "out/智能审查版.pptx"

# 纯规则模式（LLM不可用时）
python -m app.pptlint.cli --input "input.pptx" --config "configs/config.yaml" --llm off --output-ppt "out/规则版.pptx"

# 混合模式：规则+LLM（推荐）
python -m app.pptlint.cli --input "input.pptx" --config "configs/config_llm.yaml" --llm on --output-ppt "out/混合审查版.pptx"
```

### 2. GUI模式（适合单文件处理）
```bash
python -m pptlint.gui
```
- 图形化选择文件和配置
- 实时运行状态显示
- 支持LLM参数配置

### 3. WebUI模式（推荐，界面美观）
```bash
streamlit run webui/app.py
```
- 现代化Web界面
- 拖拽上传文件
- 一键运行和下载
- 保留CLI/GUI并行使用

## 📊 输出示例

### 问题报告 (Markdown)
```markdown
## 审查报告

共发现 3 项问题。

- 规则: FontSizeRule | 严重性: warning | 页: 0 | 对象: s1
  - 描述: 字号 10 < 12
  - 建议: 提升至 12pt
  - 可自动修复: 否 | 已修复: 否

- 规则: AcronymRule | 严重性: info | 页: 1 | 对象: page
  - 描述: 缩略语 API 首次出现未发现解释
  - 建议: 在首次出现后添加解释：API: <全称>
  - 可自动修复: 否 | 已修复: 否
```

### PPT标记输出
- 首页左上角显示问题汇总
- 问题对象文本标记为红色斜体下划线
- 文本末尾添加中文类别标记

## 🚧 开发状态

### 已完成模块
- ✅ PPTX解析器：完整实现
- ✅ 规则引擎：基础规则完整
- ✅ 工作流编排：核心流程完整
- ✅ 多界面支持：CLI/GUI/WebUI完整
- ✅ 配置系统：灵活配置支持
- ✅ 报告生成：Markdown格式完整
- ✅ PPT标记：问题标记完整
- ✅ **LLM智能审查：核心功能完整**

### 待完善功能
- 🔄 自动修复：基础框架已搭建，具体逻辑待实现
- 🔄 主题色检测：预留接口，待完善
- 🔄 跨页关联：缩略语解释检测范围待扩展
- 🔄 LLM提示词优化：持续优化提示词工程

### 技术债务
- 部分自动修复功能需要补充实现
- LLM集成功能需要完善
- 主题色一致性检查需要深入PPT母版解析

## 🤝 贡献指南

1. Fork 项目
2. 创建功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 开启 Pull Request

## 📄 许可证

本项目采用 MIT 许可证 - 查看 [LICENSE](LICENSE) 文件了解详情

## 📞 支持

如有问题或建议，请提交 Issue 或联系开发团队。

## 🎯 完整运行示例

### 步骤1：准备环境
```bash
# 克隆项目
git clone <your-repo-url>
cd PPTCheker

# 安装依赖
cd app
pip install -r requirements.txt

# 配置 DeepSeek API 密钥（推荐）
export DEEPSEEK_API_KEY="your_deepseek_api_key"
```

### 步骤2：运行工具
```bash
# 方式1：CLI模式（推荐批处理）
python -m pptlint.cli --input "../智能体及扣子介绍.pptx" --config "../configs/config.yaml" --report "../out/report.md" --output-ppt "../out/标记版.pptx"

# 方式2：WebUI模式（推荐单文件处理）
streamlit run webui/app.py

# 方式3：桌面GUI模式
python -m pptlint.gui
```

### 步骤3：查看结果
```bash
# 查看生成的报告
cat ../out/report.md

# 查看输出目录
ls -la ../out/
```

### 预期输出
- ✅ 成功处理PPT文件，发现问题并生成报告
- ✅ 生成带标记的PPT文件（红色斜体下划线标记问题）
- ✅ 生成Markdown格式的详细问题报告
- ✅ 首页显示问题汇总统计

