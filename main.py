import streamlit as st
import pandas as pd
import os
import time
from openai import OpenAI

# 导入模块
import logic
import auth
import styles

# 初始化
st.set_page_config(page_title="WordToWord V1.0", page_icon="📝", layout="wide")
styles.inject_css()
auth.init_db()

# Session State 初始化
if 'logged_in' not in st.session_state: st.session_state.logged_in = False
if 'user_role' not in st.session_state: st.session_state.user_role = None
if 'username' not in st.session_state: st.session_state.username = ""
if 'step' not in st.session_state: st.session_state.step = 1
if 'plan' not in st.session_state: st.session_state.plan = None
if 'template_bytes' not in st.session_state: st.session_state.template_bytes = None
# 新增：专门用来存用户原始文件名的变量，只用于显示和下载，不用于路径
if 'user_filename_display' not in st.session_state: st.session_state.user_filename_display = "template.docx"


# ================= 登录页 =================
def login_page():
    c1, c2, c3 = st.columns([1, 1, 1])
    with c2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.markdown(styles.get_logo_html(), unsafe_allow_html=True)

        tab1, tab2 = st.tabs(["🔐 登录", "📝 注册"])
        with tab1:
            with st.form("login"):
                u = st.text_input("用户名")
                p = st.text_input("密码", type="password")
                if st.form_submit_button("登录系统", type="primary", use_container_width=True):
                    role = auth.login_user(u, p)
                    if role:
                        st.session_state.logged_in = True
                        st.session_state.user_role = role
                        st.session_state.username = u
                        st.rerun()
                    else:
                        st.error("用户名或密码错误")

        with tab2:
            with st.form("reg"):
                nu = st.text_input("用户名")
                np = st.text_input("密码", type="password")
                if st.form_submit_button("注册新账号", use_container_width=True):
                    if auth.register_user(nu, np):
                        st.success("注册成功，请登录")
                    else:
                        st.error("用户名已存在")


# ================= 管理员后台 =================
def admin_page():
    st.markdown(styles.get_logo_html(), unsafe_allow_html=True)
    st.markdown("### 🛠️ 管理员控制台")

    if st.button("退出登录"):
        st.session_state.logged_in = False
        st.rerun()

    users, logs, fb = auth.get_admin_data()
    m1, m2, m3 = st.columns(3)
    m1.metric("总用户数", len(users))
    m2.metric("累计任务", len(logs))
    m3.metric("平均满意度", f"{fb['rating'].mean():.1f}" if not fb.empty else "0.0")

    st.markdown("---")
    c1, c2 = st.columns(2)
    with c1:
        st.caption("用户列表")
        st.dataframe(users, use_container_width=True, height=250)
    with c2:
        st.caption("最新反馈")
        st.dataframe(fb, use_container_width=True, height=250)

    st.caption("系统日志")
    st.dataframe(logs, use_container_width=True)


# ================= 用户工作台 =================
def user_page():
    # --- 侧边栏 ---
    with st.sidebar:
        st.title("设置")
        api_key = st.text_input("DeepSeek API Key", type="password")
        if not api_key:
            st.warning("⚠️ 请输入 API Key")
        else:
            st.success("✅ API Key 已就绪")

        st.divider()
        with st.expander("📖 V1.0 使用指南", expanded=True):
            st.markdown(styles.get_guide_html(), unsafe_allow_html=True)

        st.divider()
        with st.form("fb"):
            score = st.slider("评分", 1, 5, 5)
            txt = st.text_area("反馈")
            if st.form_submit_button("提交"):
                auth.submit_feedback(st.session_state.username, txt, score)
                st.success("已提交")

        if st.button("退出登录"):
            st.session_state.logged_in = False
            st.rerun()

    # --- 主界面 ---
    c_logo, c_user = st.columns([3, 1])
    with c_logo:
        st.markdown(styles.get_logo_html(), unsafe_allow_html=True)
    with c_user:
        st.markdown(
            f"<div style='text-align:right; color:#64748b; padding-top:20px;'>👤 {st.session_state.username}</div>",
            unsafe_allow_html=True)

    # 步骤 1: 上传
    if st.session_state.step == 1:
        st.markdown("""
        <div class="w2w-card">
            <div class="w2w-header">📂 步骤 1: 建立任务</div>
            <div class="w2w-desc">系统已升级，现在支持直接读取 PDF 格式的简历或非结构化 Word 文档。</div>
        """, unsafe_allow_html=True)

        c1, c2 = st.columns(2)
        f_old = c1.file_uploader("源文件 (简历/旧表格)", type=["docx", "pdf"], key="old")
        f_new = c2.file_uploader("目标文件 (空白模板)", type=["docx"], key="new")

        st.markdown("<br>", unsafe_allow_html=True)
        if f_old and f_new:
            if st.button("🚀 开始 AI 分析 (V1.0)", type="primary", use_container_width=True):
                if not api_key:
                    st.error("请在侧边栏输入 API Key")
                else:
                    if not os.path.exists("temp"): os.makedirs("temp")

                    # 1. 强制使用英文文件名保存到服务器
                    old_ext = os.path.splitext(f_old.name)[1]
                    p_old = os.path.join("temp", f"source_file{old_ext}")
                    p_new = os.path.join("temp", "target_template.docx")

                    # 2. 写入临时文件
                    with open(p_old, "wb") as f:
                        f.write(f_old.getbuffer())
                    with open(p_new, "wb") as f:
                        f.write(f_new.getbuffer())

                    # 3. 备份数据到 Session
                    st.session_state.template_bytes = f_new.getvalue()
                    st.session_state.user_filename_display = f_new.name

                    # ========================================================
                    # 【核心修改】在此处立刻检查文件格式！防止 Token 浪费
                    # ========================================================
                    is_valid_new, err_msg_new = logic.validate_file_format(p_new)
                    if not is_valid_new:
                        st.error(err_msg_new)
                        st.stop()  # 🛑 立即停止，不扣费

                    is_valid_old, err_msg_old = logic.validate_file_format(p_old)
                    if not is_valid_old:
                        st.error(f"源文件错误：{err_msg_old}")
                        st.stop()  # 🛑 立即停止
                    # ========================================================

                    with st.spinner("正在读取文档并构建知识图谱..."):
                        try:
                            client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
                            old_txt = logic.read_file_content(p_old)
                            new_txt = logic.read_file_content(p_new)

                            # 只有上面文件检查通过了，才会走到这一步扣费
                            plan = logic.generate_filling_plan_v2(client, old_txt, new_txt)

                            st.session_state.plan = plan
                            st.session_state.kv_df = pd.DataFrame(plan['kv'])
                            st.session_state.step = 2
                            auth.log_action(st.session_state.username, f"Analysis: {f_new.name}")
                            st.rerun()
                        except Exception as e:
                            st.error(f"处理失败: {e}")
        st.markdown("</div>", unsafe_allow_html=True)

    # 步骤 2: 审核
    elif st.session_state.step == 2:
        st.markdown("""
        <div class="w2w-card">
            <div class="w2w-header">📊 步骤 2: 数据核对</div>
            <div class="w2w-desc">AI 已从源文件中提取数据。您可以自由修改，或使用 AI 润色工具。</div>
        """, unsafe_allow_html=True)

        edited_df = st.data_editor(
            st.session_state.kv_df,
            column_config={"anchor": "字段", "val": st.column_config.TextColumn("内容", width="large"),
                           "source": "来源"},
            use_container_width=True, num_rows="dynamic", height=400
        )

        lists = st.session_state.plan.get("lists", [])
        if lists:
            st.info(f"📋 识别到 {len(lists)} 个列表，将自动扩展表格行。")
            for lst in lists:
                with st.expander(f"查看列表: {lst.get('keyword')}"):
                    st.dataframe(pd.DataFrame(lst['data'], columns=lst.get('headers')))

        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("""<div class="w2w-card"><div class="w2w-header">✨ AI 润色</div>""", unsafe_allow_html=True)
        c1, c2, c3 = st.columns([2, 2, 1])
        t_target = c1.selectbox("选择字段", edited_df['anchor'].tolist())
        t_prompt = c2.text_input("指令", placeholder="例如：扩充到200字，语气更自信")
        if c3.button("执行", use_container_width=True):
            if not api_key:
                st.error("No Key")
            else:
                client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
                idx = st.session_state.kv_df.index[st.session_state.kv_df['anchor'] == t_target].tolist()[0]
                curr = edited_df.loc[idx, 'val']
                new_val = logic.refine_text_v2(client, curr, t_prompt)
                st.session_state.kv_df.at[idx, 'val'] = new_val
                st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

        c_b1, c_b2 = st.columns(2)
        if c_b1.button("🔙 返回"):
            st.session_state.step = 1
            st.rerun()
        if c_b2.button("✅ 确认生成", type="primary"):
            st.session_state.plan['kv'] = edited_df.to_dict('records')
            st.session_state.step = 3
            st.rerun()

    # 步骤 3: 写入
    elif st.session_state.step == 3:
        st.markdown("""
        <div class="w2w-card" style="text-align:center; padding:40px;">
            <h3 style="color:#4F46E5;">⚙️ 正在写入 V1.0 文档...</h3>
            <p style="color:#6B7280;">AI 引擎正在处理格式对齐与列表克隆。</p>
        </div>
        """, unsafe_allow_html=True)

        bar = st.progress(0)

        def update_bar(p, msg):
            bar.progress(p, text=msg)
            time.sleep(0.05)

        try:
            # 文件名
            p_template = os.path.join("temp", "target_template.docx")
            p_out = os.path.join("temp", "final_result.docx")

            # 强制覆盖逻辑
            if st.session_state.get('template_bytes'):
                if not os.path.exists("temp"): os.makedirs("temp")
                with open(p_template, "wb") as f:
                    f.write(st.session_state.template_bytes)
            else:
                if not os.path.exists(p_template):
                    st.error("⚠️ 关键文件丢失（会话已过期）。请刷新页面重新上传。")
                    st.stop()

            logic.execute_word_writing_v2(
                st.session_state.plan, p_template, p_out, progress_callback=update_bar
            )

            auth.log_action(st.session_state.username, f"Completed: {st.session_state.user_filename_display}")

            st.success("处理完成！")

            # 下载逻辑
            output_name = f"WordToWord_V1.0_{st.session_state.user_filename_display}"

            with open(p_out, "rb") as f:
                st.download_button("📥 下载结果", f,
                                   file_name=output_name,
                                   mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                   type="primary", use_container_width=True)

            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("新任务"):
                st.session_state.step = 1
                st.rerun()
        except Exception as e:
            st.error(f"处理出错: {e}")
            print(f"Error detail: {str(e)}")


# ================= 路由 =================
if not st.session_state.logged_in:
    login_page()
else:
    if st.session_state.user_role == 'admin':
        admin_page()
    else:
        user_page()