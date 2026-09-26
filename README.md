# 📚 学习规划助手（Learning Plan Assistant）

基于大模型（DeepSeek）的**个性化学习规划生成工具**。用户在网页填写学历、目标、基础水平、时间等信息，系统调用大模型生成对标企业招聘要求的学习规划，并自动导出 **Word** 和 **PDF** 文档。

## ✨ 功能特性

- **Streamlit 网页表单**：收集学习信息，交互友好
- **DeepSeek 大模型接入**：OpenAI 兼容接口，流式输出（打字机效果）
- **对标企业要求**：Prompt 结合目标岗位 JD 的任职要求，拆解学习内容
- **当前时间感知**：自动以当前时间为起点安排阶段与里程碑
- **一键导出**：自动生成带层级排版的 Word 文档，并转成 PDF
- **一键下载**：页面顶部提供 PDF 下载按钮

## 🛠️ 技术栈

- Python 3.13
- Streamlit（Web 界面）
- DeepSeek API（OpenAI 兼容接口）
- python-docx（生成 Word 文档）
- pywin32 / Word COM（docx 转 pdf）

## 🚀 快速开始

### 1. 克隆项目

```bash
git clone https://github.com/hqqqz/learning-plan-assistant.git
cd learning-plan-assistant
```

### 2. 创建并激活虚拟环境

```bash
python -m venv .venv
.venv\Scripts\activate    # Windows
```

### 3. 安装依赖

```bash
pip install streamlit openai python-docx pywin32
```

### 4. 运行

```bash
streamlit run 01求职问答系统.py
```

> ⚠️ **注意**：docx 转 PDF 功能依赖本机已安装 **Microsoft Word**（通过 Word COM 转换）。未安装 Word 时，Word 文档仍会正常生成，仅 PDF 转换会失败。

### 5. 使用

1. 在侧边栏填入你的 **DeepSeek API Key**（在 [DeepSeek 开放平台](https://platform.deepseek.com) 申请）
2. 填写学习信息（学历、目标、基础水平、时间等）
3. 点击"🚀 获取学习规划"
4. 自动生成学习规划 Word + PDF，并可在页面顶部下载 PDF

## 📁 项目结构

```
learning-plan-assistant/
├── 01求职问答系统.py      # 主程序（Streamlit 应用）
├── docx_to_pdf.py         # docx 转 pdf 模块（调用 Word COM）
├── .gitignore
└── README.md
```

## 📝 说明

- 项目主要用于个人学习练手，覆盖 Streamlit、大模型 API 调用、文档自动化等技能
- 生成的 Word / PDF 保存在脚本同目录，文件名带时间戳避免覆盖
