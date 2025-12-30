<p align="center"><img src="https://img.shields.io/badge/WordToWord-V1.0-6366f1?style=for-the-badge&logo=googledocs&logoColor=white" alt="Logo"></p>

<h1 align="center">
  📝 WordToWord
</h1>

<h3 align="center">基于 DeepSeek 的智能文档迁移与自动化填表助手</h3>

<p align="center">
  告别“Ctrl+C / Ctrl+V”，让 AI 帮你搞定繁琐的表格填写。
</p>

<p align="center">
  <a href="https://word-to-word.streamlit.app/" target="_blank">
    <img src="https://static.streamlit.io/badges/streamlit_badge_black_white.svg" alt="Open in Streamlit">
  </a>
</p>

<p align="center">
  <a href="#-核心功能">核心功能</a> ·
  <a href="#-快速开始">快速开始</a> ·
  <a href="#-项目结构">项目结构</a> ·
  <a href="#-开源协议">开源协议</a>
</p>

<p align="center">
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
</p>
<br>

> **WordToWord** 是一款专为高校师生、行政人员及企业HR设计的**智能文档自动化工具**。它利用大语言模型（LLM）的深度语义理解能力，能够从非结构化的资料（PDF/Word）中提取信息，并精准填充到复杂的 Word 表格模板中。支持自动勾选、列表动态扩展及 AI 润色。

---

## 💻 在线演示 (Live Demo)

👉 **点击立即体验 V1.0 版本：[https://word-to-word.streamlit.app/](https://word-to-word.streamlit.app/)**

> *提示：为了体验完整功能，请自备 DeepSeek API Key。*

---

## 😫 你是否也经历过这样的“崩溃时刻”？

> **场景一：** 奖学金评选开始了，你要填一张新的《申请审批表》。明明你的个人简介、获奖经历、科研成果在之前的《个人简历》里都写过一遍了，但你还是得打开两个文档，**Ctrl+C，Ctrl+V，调整字体，调整格式...** 机械重复，浪费生命。

> **场景二：** 年终考核，HR 发来一张全新的 Word 模板。你需要把去年的数据搬过来，还得把“自我评价”写得更漂亮点。对着屏幕发呆半小时，憋不出一句话。

> **场景三：** 遇到一个表格，里面有几十个“是/否”的勾选框（□），你得一个个手动替换成（☑），点到手抽筋。

**WordToWord 就是为了解决这个问题而生的。**

它不是简单的复制粘贴工具，它是一个**懂你数据的 AI 填表助手**。

---

## 💡 它是怎么工作的？（一分钟看懂）

简单来说，你只需要做两步：

1.  **左手**：扔进你现有的资料（比如你的 PDF 简历、旧的 Word 申请表）。
2.  **右手**：扔进那个必须要填的空白 Word 模板。

**剩下的交给 WordToWord：**
* 它会阅读你的资料，理解你是谁。
* 它会看懂空白表格，知道每一栏该填什么。
* **它自动把内容填进去，甚至会自动帮你润色文字，把那个讨厌的勾选框打上勾！**

---

## 🚀 核心黑科技 (Core Features)

### 🧠 1. 基于知识图谱的语义理解
我们不搞低级的关键词匹配。依靠 **DeepSeek / OpenAI** 的强大能力，它能真正“读懂”文档。
* 表格里问“曾获荣誉”，它能自动从你的简历里提取“一等奖”、“优秀标兵”。
* 表格里问“生日”，它知道填“出生日期”而不是“年龄”。

### ✍️ 2. AI 智能创作与润色
* **缺失内容自动补全**：目标表格需要“自我评价”，但你简历里没写？没关系，AI 会根据你的过往经历，自动帮你写一段得体、专业的评价。
* **交互式润色**：觉得 AI 填写的太生硬？选中它，告诉 AI：“帮我改得自信一点”、“扩充到 200 字”，立刻搞定。

### 📑 3. 复杂格式完美对齐
* **列表自动克隆 (Smart List Cloning)** ✨：这是最帅的功能。如果你的简历里有 10 门课程，而表格里只有一行？系统会自动**向下加行**，并且完美保留原表格的边框、字体和格式。
* **智能勾选**：自动识别语义（Yes/No, 有/无），将 Word 中的 `□` 替换为 `☑`，治愈强迫症。

### 🛡️ 4. 安全的企业级后台
* **隐私分离**：支持用户自己输入 API Key，你的数据只经过大模型，不经过我们的服务器存储。
* **管理员看板**：后台可监控任务流、用户反馈及系统健康状态。

---

## 🛠️ 本地部署 (Developer Guide)

如果你想在自己的电脑或服务器上跑，也非常简单：

```bash
# 1. 克隆代码
git clone [https://github.com/jiahao-bot/WordToWord.git](https://github.com/jiahao-bot/WordToWord.git)
cd WordToWord

# 2. 安装依赖
pip install -r requirements.txt

# 3. 配置环境 (可选，如需本地管理员功能)
# cp .env.example .env

# 4. 启动应用
streamlit run main.py
```

访问 `http://localhost:8501` 即可。

------

## 🏗️ 项目架构

本项目采用模块化设计，各模块职责明确，便于二次开发与维护。

Plaintext

```
WordToWord/
├── main.py          # [入口] 应用主入口，负责路由分发与 Session 管理
├── logic.py         # [核心] 业务逻辑层，包含 LLM 交互、文档解析与写入算法
├── auth.py          # [安全] 鉴权模块，处理 SQLite 数据库交互、加密与权限控制
├── styles.py        # [UI] 前端样式层，包含 CSS 注入与组件渲染
├── wordtoword.db    # [数据] SQLite 数据库文件（自动生成）
└── requirements.txt # [依赖] 项目依赖清单
```

------

## 🤝 贡献指南

我们非常欢迎社区开发者参与本项目的改进。

1. **Fork 本仓库**：请将项目 Fork 到您的个人 GitHub 账户下。
2. **创建分支**：建议基于 `main` 分支创建新的功能分支。
3. **提交代码**：请确保代码风格整洁，注释清晰。
4. **发起 Pull Request**：详细说明变更内容。

------

## ⚖️ 开源协议 (License)

本项目严格遵循 **GNU Affero General Public License v3.0 (AGPL-3.0)** 开源协议。

**使用本项目即代表您同意以下条款：**

1. **开源义务**：如果您基于本项目进行修改、衍生开发，并通过网络提供服务（SaaS），您**必须**向您的用户公开完整的源代码。
2. **版权声明**：在分发或部署本项目时，必须保留原始的版权声明、许可声明及作者信息。
3. **商业限制**：本软件按“原样”提供，不提供任何形式的商业担保。任何基于本项目的闭源商业行为均被严格禁止，除非获得原作者的商业授权。

Copyright © 2025 jiahao-bot. All Rights Reserved.

------

<div align="center"> <p>Made with ❤️ by <b>jiahao-bot</b></p> <p>Powered by <b>Streamlit</b> & <b>DeepSeek</b></p> </div>

