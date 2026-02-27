# bp-analyzer
# 商业计划书 AI 分析助手

基于通义千问和 Streamlit 开发的智能商业计划书分析工具。

## 功能特点

- 📄 支持 PDF 和 PPT 格式的商业计划书上传
- 🤖 使用通义千问 AI 进行智能分析
- 📊 生成专业的多维度分析报告
- 💾 支持分析报告下载

## 在线体验

访问: [您的 Streamlit 应用链接]

## 本地运行

### 前置要求

- Python 3.8+
- 通义千问 API Key

### 安装步骤

1. 克隆仓库
```bash
git clone https://github.com/你的用户名/bp-analyzer.git
cd bp-analyzer
```

2. 安装依赖
```bash
pip install -r requirements.txt
```

3. 运行应用
```bash
streamlit run app.py
```

4. 在浏览器中打开 http://localhost:8501

## 使用说明

1. 在侧边栏输入您的通义千问 API Key
2. 上传商业计划书文件（PDF 或 PPT）
3. 点击"开始分析"按钮
4. 等待 AI 生成分析报告
5. 查看报告并可选择下载

## 获取 API Key

访问 [通义千问控制台](https://dashscope.aliyun.com/) 注册并获取 API Key。

## 技术栈

- **前端框架**: Streamlit
- **AI 模型**: 通义千问 (Qwen)
- **文件处理**: PyPDF2, python-pptx
- **API 调用**: OpenAI SDK

## 许可证

MIT License

## 联系方式

如有问题或建议，欢迎提 Issue。

