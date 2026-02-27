
import streamlit as st
import pandas as pd
from pypdf import PdfReader
from pptx import Presentation
import dashscope
import os
from io import BytesIO

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
    "上传调研要点 Excel 模板和商业计划书（PDF 或 PPT），系统会根据 Excel 结构生成项目分析报告。"
)

# ==========================
# 上传文件
# ==========================
# 上传调研要点 Excel
excel_file = st.file_uploader(
    "上传调研要点 Excel 模板（.xlsx）",
    type=["xlsx"],
    key="excel",
    help="这是调研要点列表文件，用于结构化分析。"
)

# 上传商业计划书
uploaded_file = st.file_uploader(
    "上传商业计划书（PDF 或 PPT）",
    type=["pdf", "pptx"],
    key="bp",
    help="建议文件不超过20MB"
)

# ==========================
# 读取 Excel 调研要点
# ==========================
def read_excel_survey(file):
    try:
        # 读取 Excel 的第一个 sheet
        df = pd.read_excel(file, sheet_name=0)
        # 假设结构：列0为空，列1为标题，列2为调研要点，列3为空或评价
        # 提取非空行中的调研要点（列1和列2）
        text = ""
        for index, row in df.iterrows():
            title = str(row.iloc[1]) if pd.notna(row.iloc[1]) else ""
            points = str(row.iloc[2]) if pd.notna(row.iloc[2]) else ""
            if title or points:
                text += f"\n### {title}\n{points}\n"
        return text.strip()
    except Exception as e:
        st.error(f"读取 Excel 失败: {e}")
        return ""

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
if excel_file and uploaded_file:
    # 读取 Excel 调研要点
    survey_points = read_excel_survey(excel_file)
    st.success("Excel 调研要点读取成功！")

    # 读取商业计划书
    if uploaded_file.name.endswith(".pdf"):
        bp_text = read_pdf(uploaded_file)
    else:
        bp_text = read_ppt(uploaded_file)
    st.success("商业计划书读取成功！")

    if st.button("开始分析"):
        # ==========================
        # Prompt 模板（融入 Excel 调研要点）
        # ==========================
        prompt = f"""你是一个项目评价助手。
请严格根据以下调研要点结构分析商业计划书，按模块输出分析结果。每个模块需基于要点进行详细评价，并给出'评价'（高/中/低风险或潜力）。

调研要点结构：
{survey_points}

商业计划书内容：
{bp_text}

输出格式：每个模块标题后跟分析内容，最后一行'评价：高/中/低'。禁止输出融资建议或融资规模建议。"""

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
        else:
            st.warning("未生成分析结果，请检查 API Key 或文件内容。")
else:
    st.info("请上传调研要点 Excel 和商业计划书文件。")


