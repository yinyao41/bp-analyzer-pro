
import streamlit as st
import pandas as pd
from pypdf import PdfReader
from pptx import Presentation
import dashscope
import os

# ==========================
# 获取 API Key（Streamlit Secrets）
# ==========================
dashscope.api_key = os.environ.get("DASHSCOPE_API_KEY")

# ==========================
# 页面设置
# ==========================
st.set_page_config(page_title="商业计划书智能分析", layout="wide")
st.title("商业计划书智能分析系统")
st.markdown("上传商业计划书（PDF 或 PPT），系统将**自动使用内置调研要点模板**生成结构化分析报告。")

# ==========================
# 自动加载调研要点 Excel（无需用户上传）
# ==========================
def load_survey_points():
    file_path = "调研要点.xlsx"          # ← 必须放在项目根目录
    try:
        with open(file_path, "rb") as f:
            return read_excel_survey(f)
    except FileNotFoundError:
        st.error("❌ 未找到调研要点模板文件！请将 '调研要点.xlsx' 放在项目根目录。")
        st.stop()
    except Exception as e:
        st.error(f"读取调研要点失败: {e}")
        st.stop()

# ==========================
# 读取 Excel 调研要点（保持你原来的函数）
# ==========================
def read_excel_survey(file):
    try:
        df = pd.read_excel(file, sheet_name=0)
        text = ""
        for _, row in df.iterrows():
            title = str(row.iloc[1]) if pd.notna(row.iloc[1]) else ""
            points = str(row.iloc[2]) if pd.notna(row.iloc[2]) else ""
            if title or points:
                text += f"\n### {title}\n{points}\n"
        return text.strip()
    except Exception as e:
        st.error(f"解析调研要点 Excel 失败: {e}")
        return ""

# ==========================
# 读取 PDF / PPT（保持你原来的函数）
# ==========================
def read_pdf(file):
    reader = PdfReader(file)
    text = ""
    for page in reader.pages:
        text += page.extract_text() or ""
    return text

def read_ppt(file):
    prs = Presentation(file)
    text = ""
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text"):
                text += shape.text + "\n"
    return text

# ==========================
# 加载内置调研要点（程序启动就加载）
# ==========================
survey_points = load_survey_points()
st.success("✅ 调研要点模板已自动加载")

# ==========================
# 上传商业计划书
# ==========================
uploaded_file = st.file_uploader(
    "上传商业计划书（PDF 或 PPT）",
    type=["pdf", "pptx"],
    help="建议文件不超过20MB"
)

# ==========================
# 分析流程
# ==========================
if uploaded_file:
    # 读取商业计划书
    if uploaded_file.name.endswith(".pdf"):
        bp_text = read_pdf(uploaded_file)
    else:
        bp_text = read_ppt(uploaded_file)
    
    st.success("商业计划书读取成功！")

    if st.button("开始分析"):
        prompt = f"""你是一个项目评价助手。
请严格根据以下调研要点结构分析商业计划书，按模块输出分析结果。每个模块需基于要点进行详细评价，并给出'评价'（高/中/低风险或潜力）。

调研要点结构：
{survey_points}

商业计划书内容：
{bp_text}

输出格式：每个模块标题后跟分析内容，最后一行'评价：高/中/低'。禁止输出融资建议或融资规模建议。"""

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

        if result:
            st.subheader("分析结果")
            st.write(result)
        else:
            st.warning("未生成分析结果，请检查 API Key。")
else:
    st.info("请上传商业计划书文件")
