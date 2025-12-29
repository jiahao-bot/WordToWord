<div align="center">
  <h1>📝 WordToWord</h1>
  <h3>基于 DeepSeek 的智能文档迁移与自动化填表系统</h3>

  <p>
    告别“Ctrl+C / Ctrl+V”，让 AI 帮你搞定繁琐的表格填写。
  </p>

  <a href="https://word-to-word.streamlit.app/" target="_blank">
    <img src="https://static.streamlit.io/badges/streamlit_badge_black_white.svg" alt="Open in Streamlit">
  </a>

  <br />
  <br />

  <p>
    <a href="#-核心功能">核心功能</a> ·
    <a href="#-快速开始">快速开始</a> ·
    <a href="#-项目结构">项目结构</a> ·
    <a href="#-开源协议">开源协议</a>
  </p>

  <a href="https://www.python.org/">
    <img src="https://img.shields.io/badge/Python-3.8+-blue.svg" alt="Python Version">
  </a>
  <a href="https://streamlit.io/">
    <img src="https://img.shields.io/badge/Streamlit-App-FF4B4B.svg" alt="Streamlit">
  </a>
  <a href="https://platform.openai.com/">
    <img src="https://img.shields.io/badge/AI-DeepSeek%2FOpenAI-green.svg" alt="AI Powered">
  </a>
  <a href="LICENSE">
    <img src="https://img.shields.io/badge/License-AGPL%20v3-red.svg" alt="License">
  </a>
</div>

<br />

> **WordToWord** 是一款专为高校师生、行政人员及企业HR设计的**智能文档自动化工具**。它利用大语言模型（LLM）的深度语义理解能力，能够从非结构化的简历（PDF/Word）中提取信息，并精准填充到复杂的 Word 表格模板中。支持自动勾选、列表动态扩展及 AI 润色。

---

## 💻 在线演示 (Live Demo)

👉 **点击立即体验 V1.0 版本：[https://word-to-word.streamlit.app/](https://word-to-word.streamlit.app/)**

> *提示：为了体验完整功能，请自备 DeepSeek API Key 或兼容 OpenAI 格式的 Key。*

---

## ✨ 核心功能 (Why WordToWord?)

### 🧠 1. 深度语义理解
不同于传统的关键词匹配，WordToWord 使用 **DeepSeek/OpenAI** 引擎构建知识图谱。
- 能够理解“曾获一等奖”即“获奖情况”。
- 能够根据用户画像，**自动创作**缺失的主观评价（如“思想素质”、“自我总结”）。

### 📂 2. 多格式源文件支持
- **PDF & Word**: 无论是 PDF 格式的简历，还是旧的 Word 表格，统统扔进去，系统自动解析。
- **复杂版面分析**: 内置 `pdfplumber` 和 `python-docx`，精准还原文档结构。

### ⚙️ 3. 复杂排版自动对齐
- **列表克隆 (Smart List Cloning)**: 识别到“课程”、“奖项”等列表数据时，自动在目标 Word 表格中**向下添加行**，并保持原有格式不变（字体、边框完美复刻）。
- **勾选框识别**: 自动识别“有/无”、“是/否”语义，并将 Word 中的 `□` 替换为 `☑`。

### 🎨 4. AI 交互式润色
- 觉得 AI 填写的“自我评价”太生硬？
- 内置**润色工具**，选中字段，输入指令（如：“语气更自信一点”、“扩充到200字”），AI 实时重写。

### 🛡️ 5. 企业级管理后台
- **用户鉴权**: 完整的 登录/注册 系统，数据隔离。
- **管理员看板**: 可视化查看用户数量、任务日志及用户反馈（满意度评分）。
- **安全隐私**: API Key 本地存储（或用户自行输入），数据不留痕。

---

## 🚀 快速开始

### 环境要求
- Python 3.8+
- DeepSeek API Key

### 安装步骤

1. **克隆仓库**
   ```bash
   git clone [https://github.com/jiahao-bot/WordToWord.git](https://github.com/jiahao-bot/WordToWord.git)
   cd WordToWord

1. **安装依赖**

   Bash

   ```
   pip install -r requirements.txt
   ```

2. **配置环境** (可选) 复制 `.env.example` 为 `.env` 并配置本地管理员密码（部署到 Streamlit Cloud 时请使用 Secrets 配置）。

3. **启动应用**

   Bash

   ```
   streamlit run main.py
   ```

4. **访问应用** 打开浏览器访问 `http://localhost:8501`。

------

## 🏗️ 项目结构

Plaintext

```
WordToWord/
├── main.py          # 程序入口 (路由分发、主界面容器)
├── logic.py         # 核心逻辑 (AI Prompt, Word写入算法, PDF读取)
├── auth.py          # 鉴权模块 (数据库操作, 登录注册, 管理员数据)
├── styles.py        # UI 样式 (CSS, Logo, 操作指南组件)
├── wordtoword.db    # SQLite 数据库 (自动生成)
├── requirements.txt # 依赖列表
└── temp/            # 临时文件存储区
```

------

## 🤝 贡献指南 (Contributing)

欢迎提交 Issue 和 Pull Request！

1. Fork 本仓库
2. 新建分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送分支 (`git push origin feature/AmazingFeature`)
5. 提交 Pull Request

------

## 📄 开源协议 (License)

本项目采用 **GNU Affero General Public License v3.0 (AGPL-3.0)** 协议开源。

这意味着：

1. ✅ 你可以免费使用、修改和分发本项目。
2. ❌ **严禁闭源商用**：如果你基于本项目开发了网络服务（SaaS），你必须向该服务的用户公开你的源代码。
3. ⚠️ 修改后的代码必须沿用 AGPL-3.0 协议。

------

<div align="center"> <p>Made with ❤️ by <b>jiahao-bot</b></p> <p>Powered by <b>Streamlit</b> & <b>DeepSeek</b></p> </div>
