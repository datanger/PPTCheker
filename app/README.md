# PPT 格式审查工具

本项目依据《docs/需求文档_PPT格式审查工具.md》与《docs/架构设计_PPT格式审查工具.md》实现。

## 快速开始
```bash
pip install -r requirements.txt
python -m pptlint.cli --input "./示例.pptx" --config "./configs/config.yaml" --report "./out/report.md"
```

## 功能概述
- 检查：字体/字号、英文缩略语解释、整体色调一致、颜色数量；
- 可选自动修复：字体/字号/主题色映射；
- 建议：日文流畅性与逻辑连贯通过LLM（可选）给出建议；
- 输出：Markdown/HTML 报告。

## 目录
- pptlint/: 工具源码
- configs/: 配置与阈值
- dicts/: 术语词库
- docs/: 需求与架构文档

## GUI 运行（可选）
```bash
python -m pptlint.gui
```
- 在界面中选择：输入PPT/目录、审查需求文档、配置文件、运行模式（review/edit）、输出PPT路径；
- 在“大模型配置”处设置 provider/model/api_key（endpoint 可留空按 model 自动推断，或手动覆盖）。

## WebUI 运行（更美观，Streamlit）
```bash
streamlit run webui/app.py
```
- 上传PPT与可选的审查需求文档，设置模式与LLM配置，点击“运行”；
- 运行完成后可直接下载输出PPT；
- 保留CLI/GUI并行使用，互不影响。

