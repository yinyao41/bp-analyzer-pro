import streamlit as st
from pypdf import PdfReader
from pptx import Presentation
import dashscope
import os

# ==========================
# 获取 API Key（Streamlit Secrets）
# ==========================
# 在 Streamlit Cloud: Settings → Secrets
# 添加一行：DASHSCOPE_API_KEY=你的通义千问API_KEY
dashscope.api_key = os.environ.get("DASHSCOPE_API_KEY")

# ==========================
# Streamlit 页面设置
# ==========================
st.set_page_config(page_title="商业计划书智能分析", layout="wide")
st.title("商业计划书智能分析系统")
st.markdown(
    "上传你的商业计划书（PDF 或 PPT），系统会生成项目分析报告。"
)

# ==========================
# 上传文件
# ==========================
uploaded_file = st.file_uploader(
    "上传商业计划书（PDF 或 PPT）", type=["pdf", "pptx"], help="建议文件不超过20MB"
)

# ==========================
# 读取 PDF
# ==========================
def read_pdf(file):
    reader = PdfReader(file)
    text = ""
    for page in reader.pages:
        text += page.extract_text() or ""
    return text

# ==========================
# 读取 PPT
# ==========================
def read_ppt(file):
    prs = Presentation(file)
    text = ""
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text += shape.text + "\n"
    return text

# ==========================
# 点击分析按钮
# ==========================
if uploaded_file:
    # 判断文件类型
    if uploaded_file.name.endswith(".pdf"):
        bp_text = read_pdf(uploaded_file)
    else:
        bp_text = read_ppt(uploaded_file)

    st.success("文件读取成功！")

    if st.button("开始分析"):
        # ==========================
        # Prompt 模板
        # ==========================
        prompt = f"""
你是一个项目评价助手。

请严格分析以下商业计划书，并按模块输出分析结果：
- 产品技术
- 市场
- 行业竞争情况
- 核心团队构成
- 财务指标
- 公司架构合规情况
- 融资规模及资金用途

禁止输出融资建议或融资规模建议。

商业计划书内容：
{bp_text}
"""

        # ==========================
        # 调用通义千问
        # ==========================
        with st.spinner("分析中...请稍等"):
            try:
                response = dashscope.Generation.call(
                    model="qwen-max",
                    prompt=prompt
                )
                result = response.output.text
            except Exception as e:
                st.error(f"分析失败: {e}")
                result = ""

        # ==========================
        # 显示分析结果
        # ==========================
        if result:
            st.subheader("分析结果")
            st.write(result)
