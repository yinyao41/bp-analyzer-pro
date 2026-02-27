import streamlit as st
from pypdf import PdfReader
from pptx import Presentation
import dashscope
import os
import pandas as pd
import requests
import json

# ==========================
# 配置
# ==========================
st.set_page_config(page_title="商业计划书智能分析", layout="wide")

# 获取 API Keys
dashscope.api_key = os.environ.get("DASHSCOPE_API_KEY")
# 飞书相关配置（需要在 Streamlit Secrets 中配置）
FEISHU_APP_ID = os.environ.get("FEISHU_APP_ID", "")
FEISHU_APP_SECRET = os.environ.get("FEISHU_APP_SECRET", "")

# ==========================
# 页面标题
# ==========================
st.title("📊 商业计划书智能分析系统（飞书云文档集成版）")
st.markdown("上传商业计划书（PDF 或 PPT），系统会按照飞书 Excel 模板格式生成分析报告。")

# ==========================
# 侧边栏配置
# ==========================
with st.sidebar:
    st.header("⚙️ 配置")
    
    # 飞书文档配置
    st.subheader("飞书云文档配置")
    feishu_token = st.text_input(
        "飞书文档 Token",
        help="飞书云文档的 Token，格式如：xxxxx-xxxxxxxxxxxxxxx"
    )
    
    sheet_id = st.text_input(
        "工作表 ID",
        value="",
        help="要读取的工作表 ID，如：sheet1"
    )
    
    use_template = st.checkbox("使用飞书模板格式", value=True)
    
    st.markdown("---")
    st.markdown("""
    ### 📝 使用说明
    1. 在飞书创建 Excel 分析模板
    2. 复制文档 Token 和工作表 ID
    3. 上传商业计划书文件
    4. 点击"开始分析"
    
    ### 🔑 获取飞书凭证
    - [飞书开放平台](https://open.feishu.cn/)
    - 创建企业自建应用
    - 获取 App ID 和 App Secret
    """)

# ==========================
# 飞书 API 相关函数
# ==========================

def get_feishu_tenant_access_token(app_id, app_secret):
    """获取飞书 tenant_access_token"""
    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal"
    headers = {"Content-Type": "application/json"}
    data = {
        "app_id": app_id,
        "app_secret": app_secret
    }
    
    try:
        response = requests.post(url, headers=headers, json=data)
        result = response.json()
        if result.get("code") == 0:
            return result.get("tenant_access_token")
        else:
            st.error(f"获取 token 失败: {result.get('msg')}")
            return None
    except Exception as e:
        st.error(f"请求失败: {str(e)}")
        return None

def read_feishu_excel(token, sheet_token, range_str="A1:Z100"):
    """读取飞书云文档 Excel 内容"""
    if not FEISHU_APP_ID or not FEISHU_APP_SECRET:
        st.warning("未配置飞书 App ID 和 App Secret，请在 Streamlit Secrets 中配置")
        return None
    
    # 获取 access_token
    access_token = get_feishu_tenant_access_token(FEISHU_APP_ID, FEISHU_APP_SECRET)
    if not access_token:
        return None
    
    # 读取工作表数据
    url = f"https://open.feishu.cn/open-apis/sheets/v2/spreadsheets/{token}/values/{sheet_token}!{range_str}"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json"
    }
    
    try:
        response = requests.get(url, headers=headers)
        result = response.json()
        
        if result.get("code") == 0:
            values = result.get("data", {}).get("values", [])
            if values:
                # 转换为 DataFrame
                df = pd.DataFrame(values[1:], columns=values[0])
                return df
            else:
                st.warning("工作表为空")
                return None
        else:
            st.error(f"读取失败: {result.get('msg')}")
            return None
    except Exception as e:
        st.error(f"读取飞书文档失败: {str(e)}")
        return None

def parse_template_structure(df):
    """解析模板结构，提取分析维度"""
    if df is None or df.empty:
        return None
    
    template_structure = {}
    current_section = None
    
    for index, row in df.iterrows():
        # 假设第一列是分析维度名称，第二列是说明/要求
        dimension = str(row.iloc[0]).strip() if len(row) > 0 else ""
        description = str(row.iloc[1]).strip() if len(row) > 1 else ""
        
        if dimension and dimension != "nan":
            # 判断是否是主标题（通常全大写或特殊标记）
            if dimension.isupper() or dimension.startswith("#"):
                current_section = dimension.strip("#").strip()
                template_structure[current_section] = {
                    "description": description,
                    "items": []
                }
            elif current_section:
                template_structure[current_section]["items"].append({
                    "name": dimension,
                    "description": description
                })
    
    return template_structure

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
# 根据模板生成 Prompt
# ==========================
def generate_template_prompt(bp_text, template_structure):
    """根据飞书模板结构生成分析 Prompt"""
    
    if not template_structure:
        # 默认模板
        return f"""
你是一个专业的项目评价助手。
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
    
    # 根据模板生成 Prompt
    prompt_parts = [
        "你是一个专业的项目评价助手。",
        "请严格按照以下模板结构分析商业计划书，对每个维度进行详细分析：",
        ""
    ]
    
    for section, content in template_structure.items():
        prompt_parts.append(f"## {section}")
        if content.get("description"):
            prompt_parts.append(f"说明: {content['description']}")
        
        for item in content.get("items", []):
            prompt_parts.append(f"- {item['name']}")
            if item.get("description"):
                prompt_parts.append(f"  要求: {item['description']}")
        prompt_parts.append("")
    
    prompt_parts.extend([
        "请确保：",
        "1. 严格按照上述模板结构输出",
        "2. 每个维度都要有具体分析内容",
        "3. 禁止输出融资建议或融资规模建议",
        "4. 如果商业计划书中缺少某些信息，请标注\"信息缺失\"",
        "",
        f"商业计划书内容：\n{bp_text}"
    ])
    
    return "\n".join(prompt_parts)

def format_output_by_template(analysis_result, template_structure):
    """按照模板格式化输出结果为表格"""
    if not template_structure:
        return analysis_result
    
    # 尝试将分析结果按模板结构整理成表格
    output_data = []
    
    for section, content in template_structure.items():
        output_data.append({
            "分析维度": f"【{section}】",
            "分析内容": content.get("description", "")
        })
        
        for item in content.get("items", []):
            # 这里可以尝试从 analysis_result 中提取对应内容
            # 简化版本：直接显示项目名
            output_data.append({
                "分析维度": f"  └─ {item['name']}",
                "分析内容": item.get("description", "待分析")
            })
    
    return pd.DataFrame(output_data)

# ==========================
# 主界面
# ==========================

# 显示模板信息
if use_template and feishu_token and sheet_id:
    with st.expander("📋 查看飞书模板结构", expanded=False):
        if st.button("读取飞书模板"):
            with st.spinner("正在读取飞书云文档..."):
                template_df = read_feishu_excel(feishu_token, sheet_id)
                
                if template_df is not None:
                    st.success("✅ 模板读取成功！")
                    st.dataframe(template_df, use_container_width=True)
                    
                    # 解析模板结构
                    template_structure = parse_template_structure(template_df)
                    if template_structure:
                        st.session_state.template_structure = template_structure
                        st.json(template_structure)
                    else:
                        st.warning("无法解析模板结构")
                else:
                    st.error("模板读取失败")

# 文件上传
uploaded_file = st.file_uploader(
    "📁 上传商业计划书（PDF 或 PPT）",
    type=["pdf", "pptx"],
    help="建议文件不超过20MB"
)

# ==========================
# 分析流程
# ==========================
if uploaded_file:
    # 读取文件
    if uploaded_file.name.endswith(".pdf"):
        bp_text = read_pdf(uploaded_file)
    else:
        bp_text = read_ppt(uploaded_file)
    
    st.success(f"✅ 文件读取成功！共提取 {len(bp_text)} 个字符")
    
    # 显示提取的文本预览
    with st.expander("📄 查看提取的文本（前 500 字）"):
        st.text(bp_text[:500] + "..." if len(bp_text) > 500 else bp_text)
    
    # 分析按钮
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("🚀 开始分析", type="primary", use_container_width=True):
            # 获取模板结构
            template_structure = st.session_state.get("template_structure", None)
            
            # 生成 Prompt
            if use_template and template_structure:
                prompt = generate_template_prompt(bp_text, template_structure)
                st.info("📋 使用飞书模板格式进行分析")
            else:
                prompt = generate_template_prompt(bp_text, None)
                st.info("📋 使用默认格式进行分析")
            
            # 调用通义千问
            with st.spinner("🤖 AI 正在深度分析中，请稍候..."):
                try:
                    response = dashscope.Generation.call(
                        model="qwen-max",
                        prompt=prompt
                    )
                    result = response.output.text
                    
                    if result:
                        st.success("✅ 分析完成！")
                        
                        # 显示结果
                        st.markdown("---")
                        st.subheader("📊 分析结果")
                        
                        # 如果使用模板，尝试格式化输出
                        if use_template and template_structure:
                            # 显示原始分析结果
                            with st.expander("📝 查看详细分析", expanded=True):
                                st.markdown(result)
                            
                            # 显示模板对照表
                            st.markdown("### 📋 按模板结构整理")
                            output_df = format_output_by_template(result, template_structure)
                            st.dataframe(output_df, use_container_width=True)
                        else:
                            st.markdown(result)
                        
                        # 下载按钮
                        st.markdown("---")
                        col1, col2, col3 = st.columns(3)
                        
                        with col1:
                            st.download_button(
                                label="📥 下载分析报告（TXT）",
                                data=result,
                                file_name=f"分析报告_{uploaded_file.name}.txt",
                                mime="text/plain"
                            )
                        
                        with col2:
                            if use_template and template_structure:
                                csv_data = output_df.to_csv(index=False, encoding='utf-8-sig')
                                st.download_button(
                                    label="📥 下载结构化报告（CSV）",
                                    data=csv_data,
                                    file_name=f"结构化报告_{uploaded_file.name}.csv",
                                    mime="text/csv"
                                )
                    else:
                        st.error("分析结果为空，请重试")
                        
                except Exception as e:
                    st.error(f"❌ 分析失败: {str(e)}")
                    st.exception(e)

# ==========================
# 使用说明
# ==========================
with st.expander("📖 使用指南"):
    st.markdown("""
    ## 🎯 功能说明
    
    ### 1. 基础模式（不使用飞书模板）
    - 直接上传商业计划书
    - 使用默认分析模板
    - 生成标准分析报告
    
    ### 2. 飞书模板模式（推荐）
    - 在飞书创建自定义 Excel 分析模板
    - 配置飞书 App ID 和 App Secret
    - 输入文档 Token 和工作表 ID
    - 系统按照您的模板格式进行分析
    
    ## 📋 飞书模板格式要求
    
    Excel 第一列为分析维度名称，第二列为说明/要求，例如：
    
    | 分析维度 | 说明/要求 |
    |---------|-----------|
    | # 产品技术 | 评估产品技术创新性 |
    | 产品描述 | 详细描述产品功能和特点 |
    | 技术壁垒 | 分析技术门槛和壁垒 |
    | 专利情况 | 说明专利申请和拥有情况 |
    | # 市场分析 | 评估市场规模和机会 |
    | 市场规模 | 目标市场的规模和增长率 |
    | 目标客户 | 清晰定义目标客户群体 |
    
    **注意**：以 # 开头的为主标题
    
    ## 🔑 配置飞书凭证
    
    ### Streamlit Cloud 部署
    在 Streamlit Cloud 设置中添加 Secrets：
    ```toml
    DASHSCOPE_API_KEY = "你的通义千问API密钥"
    FEISHU_APP_ID = "你的飞书AppID"
    FEISHU_APP_SECRET = "你的飞书AppSecret"
    ```
    
    ### 本地运行
    创建 `.streamlit/secrets.toml` 文件：
    ```toml
    DASHSCOPE_API_KEY = "你的通义千问API密钥"
    FEISHU_APP_ID = "你的飞书AppID"
    FEISHU_APP_SECRET = "你的飞书AppSecret"
    ```
    
    ## 📝 获取飞书文档 Token
    
    1. 打开飞书云文档
    2. 文档 URL 中的 Token：
       `https://xxx.feishu.cn/sheets/shtcnxxxxxx`
       其中 `shtcnxxxxxx` 就是文档 Token
    3. 工作表 ID 通常为 `sheet1`, `sheet2` 等
    """)

# 页脚
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: gray;'>
    <p>商业计划书智能分析系统（飞书云文档集成版） | Powered by 通义千问 & Streamlit</p>
    <p style='font-size: 0.9em;'>⚠️ 本工具仅供参考，投资决策请综合多方信息判断</p>
</div>
""", unsafe_allow_html=True)


