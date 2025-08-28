from app.pptlint.workflow import run_review_workflow
from app.pptlint.config import ToolConfig

# 提供原始PPTX文件路径以支持生成标记PPT
run_review_workflow("parsing_result.json", ToolConfig(), "output.pptx", None, "example2.pptx")