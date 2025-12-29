import streamlit as st


def get_logo_html():
    # 保持顶格写法，防止代码块问题
    return """<div style="display:flex; align-items:center; gap:16px; margin-bottom:24px; padding-bottom:16px; border-bottom:1px solid #e2e8f0;">
<div style="width:48px; height:48px; flex-shrink:0;">
<svg viewBox="0 0 48 48" fill="none" xmlns="http://www.w3.org/2000/svg">
<rect width="48" height="48" rx="12" fill="url(#paint0_linear)"/>
<path d="M11 15L17 33L24 18L31 33L37 15" stroke="white" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"/>
<defs>
<linearGradient id="paint0_linear" x1="0" y1="0" x2="48" y2="48" gradientUnits="userSpaceOnUse">
<stop stop-color="#6366F1"/>
<stop offset="1" stop-color="#A855F7"/>
</linearGradient>
</defs>
</svg>
</div>
<div>
<div style="display:flex; align-items:center; gap:8px;">
<span style="font-family:'Inter', sans-serif; font-size:1.6rem; font-weight:800; color:#1e293b; letter-spacing:-0.5px; line-height:1;">WordToWord</span>
</div>
<div style="display:flex; align-items:center; gap:8px; margin-top:4px;">
<span style="font-size:0.8rem; color:#64748b; font-weight:500;">智能文档自动化系统</span>
<span style="background:#EFF6FF; color:#4F46E5; padding:2px 6px; border-radius:4px; font-size:0.7rem; font-weight:700; border:1px solid #DBEAFE;">V1.0</span>
</div>
</div>
</div>"""


def get_guide_html():
    # 保持顶格写法
    return """<div style="background:#f8fafc; border-radius:12px; padding:20px; border:1px solid #e2e8f0;">
<div style="font-weight:700; color:#334155; margin-bottom:15px; font-size:1.1rem;">📖 操作指南</div>
<div style="display:flex; gap:12px; margin-bottom:15px;">
<div style="background:#e0e7ff; width:30px; height:30px; border-radius:50%; display:flex; align-items:center; justify-content:center; color:#4f46e5; flex-shrink:0;">1</div>
<div>
<strong style="color:#1e293b; display:block;">准备文件</strong>
<span style="color:#64748b; font-size:0.85rem;">上传含有数据的“旧简历”(.pdf/.docx) 和空白的“新模板”(.docx)。</span>
</div>
</div>
<div style="display:flex; gap:12px; margin-bottom:15px;">
<div style="background:#e0e7ff; width:30px; height:30px; border-radius:50%; display:flex; align-items:center; justify-content:center; color:#4f46e5; flex-shrink:0;">2</div>
<div>
<strong style="color:#1e293b; display:block;">AI 神经分析</strong>
<span style="color:#64748b; font-size:0.85rem;">点击按钮，AI 将自动提取 KV、列表并识别勾选框。</span>
</div>
</div>
<div style="display:flex; gap:12px; margin-bottom:15px;">
<div style="background:#e0e7ff; width:30px; height:30px; border-radius:50%; display:flex; align-items:center; justify-content:center; color:#4f46e5; flex-shrink:0;">3</div>
<div>
<strong style="color:#1e293b; display:block;">审核与润色</strong>
<span style="color:#64748b; font-size:0.85rem;">您可以双击修改内容，或让 AI 重新润色某段文字。</span>
</div>
</div>
<div style="display:flex; gap:12px;">
<div style="background:#e0e7ff; width:30px; height:30px; border-radius:50%; display:flex; align-items:center; justify-content:center; color:#4f46e5; flex-shrink:0;">4</div>
<div>
<strong style="color:#1e293b; display:block;">一键写入</strong>
<span style="color:#64748b; font-size:0.85rem;">确认无误后，生成并下载完美排版的 Word 文档。</span>
</div>
</div>
</div>"""


def inject_css():
    st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;800&family=Noto+Sans+SC:wght@400;500;700&display=swap');

        :root {
            --primary: #4F46E5;
            --bg-body: #F9FAFB;
        }

        .stApp {
            font-family: 'Inter', 'Noto Sans SC', sans-serif;
            background-color: var(--bg-body);
        }

        /* 【关键修改】padding-top 改为 6rem，解决被顶部菜单栏遮挡的问题 */
        .block-container {
            max-width: 1200px;
            padding-top: 6rem;
            padding-bottom: 5rem;
        }

        .w2w-card {
            background: white;
            border-radius: 16px;
            padding: 24px;
            border: 1px solid #E5E7EB;
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
            margin-bottom: 24px;
        }

        .w2w-header {
            font-size: 1.15rem;
            font-weight: 700;
            color: #1F2937;
            margin-bottom: 8px;
            display: flex;
            align-items: center;
            gap: 10px;
        }

        .w2w-desc {
            font-size: 0.95rem;
            color: #6B7280;
            margin-bottom: 16px;
        }

        .stButton button {
            border-radius: 10px !important;
            font-weight: 600 !important;
            height: 44px !important;
        }
    </style>
    """, unsafe_allow_html=True)