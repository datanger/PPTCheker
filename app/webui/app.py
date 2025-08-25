import os
import sys
from pathlib import Path
ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))
import tempfile
import streamlit as st

from pptlint.config import load_config, ToolConfig
from pptlint.user_req import parse_user_requirements
from pptlint.workflow import run_review_workflow, run_edit_workflow
from pptlint.llm import LLMClient


st.set_page_config(page_title="PPT 审查工具", page_icon="📊", layout="centered")

# 轻量美化
st.markdown(
    """
    <style>
      .stApp { background: #fafafa; }
      .section-card { background: #ffffff; padding: 16px 18px; border-radius: 10px; border: 1px solid #eee; }
      .section-title { font-size: 1.05rem; font-weight: 600; margin: 0 0 8px 0; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("📊 PPT 审查工具 · WebUI")
st.caption("审查/编辑两种模式 · LLM 可配置 · 一键运行")

with st.form("cfg_form"):
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">基础设置</div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        input_file = st.file_uploader("输入PPT(.pptx)", type=["pptx"], accept_multiple_files=False)
        user_req = st.file_uploader("审查需求文档(.md/.yaml)", type=["md", "yaml", "yml"], accept_multiple_files=False)
        cfg_path = st.text_input("配置文件路径", value="configs/config.yaml")
    with col2:
        mode = st.selectbox("运行模式", options=["review", "edit"], index=0)
        out_name = st.text_input("输出PPT文件名", value="标记版.pptx")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-card" style="margin-top:12px">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">LLM 设置</div>', unsafe_allow_html=True)
    col3, col4 = st.columns(2)
    with col3:
        provider = st.text_input("Provider", value="deepseek")
        model = st.text_input("Model", value="deepseek-chat")
    with col4:
        endpoint = st.text_input("Endpoint(可空自动推断)", value="")
        api_key = st.text_input("API Key", value=os.getenv("LLM_API_KEY", ""), type="password")
    st.markdown('</div>', unsafe_allow_html=True)

    submitted = st.form_submit_button("🚀 运行")

if submitted:
    if not input_file or not out_name:
        st.error("请输入输入PPT与输出文件名")
    else:
        with st.spinner("处理中，请稍候..."):
            # 将上传的输入与需求文档落盘到临时文件
            tmpdir = tempfile.mkdtemp()
            in_path = os.path.join(tmpdir, input_file.name)
            with open(in_path, "wb") as f:
                f.write(input_file.read())
            user_req_path = None
            if user_req is not None:
                user_req_path = os.path.join(tmpdir, user_req.name)
                with open(user_req_path, "wb") as f:
                    f.write(user_req.read())

            # 加载配置
            cfg: ToolConfig = load_config(cfg_path)
            if user_req_path:
                cfg = parse_user_requirements(user_req_path, cfg)

            # LLM 客户端
            ep = endpoint or None
            ak = api_key or None
            llm = LLMClient(endpoint=ep, api_key=ak, model=model)

            out_dir = os.path.join("out")
            os.makedirs(out_dir, exist_ok=True)
            out_path = os.path.join(out_dir, out_name)

            if mode == "review":
                res = run_review_workflow(in_path, cfg, out_path, llm)
            else:
                res = run_edit_workflow(in_path, cfg, out_path, llm)

            st.success(f"完成：问题 {len(res.issues)}，输出：{out_path}")
            st.download_button("下载输出PPT", data=open(out_path, "rb").read(), file_name=os.path.basename(out_path))

st.markdown("---")
st.caption("亦可使用 CLI 或桌面GUI 运行")


