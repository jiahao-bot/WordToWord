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

# Session State
if 'logged_in' not in st.session_state: st.session_state.logged_in = False
if 'user_role' not in st.session_state: st.session_state.user_role = None
if 'username' not in st.session_state: st.session_state.username = ""
if 'step' not in st.session_state: st.session_state.step = 1
if 'plan' not in st.session_state: st.session_state.plan = None
if 'template_bytes' not in st.session_state: st.session_state.template_bytes = None
if 'user_filename_display' not in st.session_state: st.session_state.user_filename_display = "template.docx"
# 新增：用于存储当前使用的源数据文本（用于展示）
if 'source_text_display' not in st.session_state: st.session_state.source_text_display = ""
if 'last_usage' not in st.session_state: st.session_state.last_usage = None
if 'preview_bytes' not in st.session_state: st.session_state.preview_bytes = None
if 'preview_data' not in st.session_state: st.session_state.preview_data = None


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
    st.dataframe(logs, use_container_width=True)


# ================= 用户工作台 =================
def user_page():
    # --- 【新增】初始化一个固定的档案名，防止每次刷新都变 ---
    if 'auto_profile_name' not in st.session_state:
        st.session_state.auto_profile_name = f"{st.session_state.username}的简历_{int(time.time())}"
    # --- 1. 侧边栏 (记忆功能核心) ---
    with st.sidebar:
        st.title("设置")
        # 自动加载 API Key
        saved_key = auth.get_user_apikey(st.session_state.username)
        api_key = st.text_input("DeepSeek API Key", value=saved_key, type="password")
        raw_model_config = auth.get_user_model_config(st.session_state.username)
        model_settings = logic.normalize_model_settings({
            "base_url": raw_model_config.get("base_url"),
            "model": raw_model_config.get("model_name"),
            "temperature": raw_model_config.get("temperature"),
            "top_p": raw_model_config.get("top_p"),
            "max_tokens": raw_model_config.get("max_tokens"),
            "frequency_penalty": raw_model_config.get("frequency_penalty"),
            "presence_penalty": raw_model_config.get("presence_penalty"),
        })

        # 如果 Key 变了，自动保存
        if api_key != saved_key and api_key:
            auth.save_user_apikey(st.session_state.username, api_key)
            st.toast("✅ API Key 已自动保存")

        if not api_key: st.warning("⚠️ 请输入 API Key")

        st.divider()
        with st.expander("🤖 模型设置", expanded=False):
            base_url = st.text_input("API Base URL", value=model_settings["base_url"])
            model_name = st.text_input("模型名称", value=model_settings["model"])
            temperature = st.slider("temperature", 0.0, 1.0, value=float(model_settings["temperature"]), step=0.05)
            top_p = st.slider("top_p", 0.0, 1.0, value=float(model_settings["top_p"]), step=0.05)
            max_tokens_input = st.text_input(
                "max_tokens (留空表示不限制)",
                value="" if model_settings["max_tokens"] is None else str(int(model_settings["max_tokens"])),
            )
            frequency_penalty = st.slider("frequency_penalty", -2.0, 2.0, value=float(model_settings["frequency_penalty"]), step=0.1)
            presence_penalty = st.slider("presence_penalty", -2.0, 2.0, value=float(model_settings["presence_penalty"]), step=0.1)
            max_tokens_value = None
            if max_tokens_input.strip():
                if max_tokens_input.strip().isdigit():
                    max_tokens_value = int(max_tokens_input.strip())
                else:
                    st.warning("max_tokens 必须为整数")

            new_model_config = {
                "base_url": base_url.strip(),
                "model_name": model_name.strip(),
                "temperature": float(temperature),
                "top_p": float(top_p),
                "max_tokens": max_tokens_value,
                "frequency_penalty": float(frequency_penalty),
                "presence_penalty": float(presence_penalty),
            }
            if new_model_config != raw_model_config:
                auth.save_user_model_config(st.session_state.username, new_model_config)
                st.toast("✅ 模型配置已保存")

            model_settings = logic.normalize_model_settings({
                "base_url": new_model_config["base_url"],
                "model": new_model_config["model_name"],
                "temperature": new_model_config["temperature"],
                "top_p": new_model_config["top_p"],
                "max_tokens": new_model_config["max_tokens"],
                "frequency_penalty": new_model_config["frequency_penalty"],
                "presence_penalty": new_model_config["presence_penalty"],
            })

            if st.session_state.last_usage:
                usage = st.session_state.last_usage
                st.caption(
                    f"本次消耗 tokens：prompt={usage.get('prompt_tokens')} / completion={usage.get('completion_tokens')} / total={usage.get('total_tokens')}"
                )

        with st.expander("📖 V1.0 使用指南", expanded=False):
            st.markdown(styles.get_guide_html(), unsafe_allow_html=True)

        # 档案管理
        st.divider()
        st.caption("📚 我的档案库")
        profiles_df = auth.get_user_profiles(st.session_state.username)
        if not profiles_df.empty:
            st.dataframe(profiles_df[['profile_name', 'created_at']], hide_index=True)
        else:
            st.info("暂无存档，上传文件后可保存。")

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

    # ================== 步骤 1: 建立任务 (档案/上传) ==================
    if st.session_state.step == 1:
        st.markdown(
            """<div class="w2w-card"><div class="w2w-header">📂 步骤 1: 建立任务</div><div class="w2w-desc">选择已有档案，或上传新文件。</div>""",
            unsafe_allow_html=True)

        # 核心升级：Tab页切换
        t1, t2 = st.tabs(["📤 上传新简历", "🗂️ 从档案库选择"])

        p_old_text = None  # 用于存储最终选定的源文本

        # 方式 A: 上传
        with t1:
            c1, c2 = st.columns(2)
            f_old = c1.file_uploader("源文件 (简历/旧表格)", type=["docx", "pdf"], key="old")
            f_new = c2.file_uploader("目标文件 (空白模板)", type=["docx"], key="new")

            # 立即检测 (UI 交互改进)
            if f_new:
                if not os.path.exists("temp"): os.makedirs("temp")
                temp_check_path = os.path.join("temp", "check_template.docx")
                with open(temp_check_path, "wb") as f:
                    f.write(f_new.getbuffer())

                valid, msg = logic.validate_file_format(temp_check_path)
                if not valid:
                    st.error(msg)
                    st.stop()  # 🛑 立即停止，不让用户点开始

            # 档案保存选项
            save_profile = st.checkbox("💾 将此源文件存为档案 (方便下次直接用)", value=True)
            # 【修改】使用 session_state 中的固定名字作为 value
            profile_name = st.text_input("档案名称",
                                         value=st.session_state.auto_profile_name,
                                         key="input_profile_name") if save_profile else ""

            # 【新增】如果不加这一行，用户修改后的值可能无法即时回写到 auto_profile_name 用于下一次刷新
            if save_profile and profile_name:
                st.session_state.auto_profile_name = profile_name

        # 方式 B: 档案
        with t2:
            profiles = auth.get_user_profiles(st.session_state.username)
            selected_profile_name = st.selectbox("选择档案",
                                                 profiles['profile_name'].tolist() if not profiles.empty else [])
            f_new_archive = st.file_uploader("目标文件 (空白模板)", type=["docx"], key="new_archive")
            if not profiles.empty and selected_profile_name:
                p_old_text = profiles[profiles['profile_name'] == selected_profile_name]['content_text'].values[0]
                st.info(f"✅ 已加载档案内容 (长度: {len(p_old_text)} 字)")

        st.markdown("<br>", unsafe_allow_html=True)

        # 统一处理开始逻辑
        start_btn = st.button("🚀 开始 AI 分析 (V1.0)", type="primary", use_container_width=True)

        if start_btn:
            if not api_key:
                st.error("请先在左侧输入 API Key")
                st.stop()

            # 确定源数据来源
            final_old_txt = ""
            final_new_path = ""
            st.session_state.preview_bytes = None
            st.session_state.preview_data = None
            st.session_state.last_usage = None

            # 路径 1: 新上传
            if f_old and f_new:
                if not os.path.exists("temp"): os.makedirs("temp")
                # 保存源文件
                old_ext = os.path.splitext(f_old.name)[1]
                p_old_path = os.path.join("temp", f"source_file{old_ext}")
                with open(p_old_path, "wb") as f:
                    f.write(f_old.getbuffer())

                # 保存目标文件 (英文名)
                final_new_path = os.path.join("temp", "target_template.docx")
                with open(final_new_path, "wb") as f:
                    f.write(f_new.getbuffer())

                # 读取内容
                final_old_txt = logic.read_file_content(p_old_path)

                # 存档案
                if save_profile and profile_name:
                    auth.save_profile(st.session_state.username, profile_name, final_old_txt)
                    st.toast("✅ 档案已保存！")

                # 存Session
                st.session_state.template_bytes = f_new.getvalue()
                st.session_state.user_filename_display = f_new.name

            # 路径 2: 用档案
            elif p_old_text and (f_new or f_new_archive):
                final_file = f_new if f_new else f_new_archive
                if not os.path.exists("temp"): os.makedirs("temp")
                final_new_path = os.path.join("temp", "target_template.docx")
                with open(final_new_path, "wb") as f:
                    f.write(final_file.getbuffer())

                final_old_txt = p_old_text
                st.session_state.template_bytes = final_file.getvalue()
                st.session_state.user_filename_display = final_file.name
            else:
                st.error("请上传文件或选择档案")
                st.stop()

            # 开始分析
            with st.spinner("正在读取文档并构建知识图谱..."):
                try:
                    # 再次预检目标文件
                    valid, msg = logic.validate_file_format(final_new_path)
                    if not valid:
                        st.error(msg)
                        st.stop()

                    new_txt = logic.read_file_content(final_new_path)
                    st.session_state.source_text_display = final_old_txt  # 存下来给用户看

                    client = OpenAI(api_key=api_key, base_url=model_settings["base_url"])
                    plan = logic.generate_filling_plan_v2(client, final_old_txt, new_txt, model_settings=model_settings)
                    plan = logic.adjust_plan_for_template(plan, final_new_path)

                    st.session_state.plan = plan
                    st.session_state.kv_df = pd.DataFrame(plan['kv'])
                    st.session_state.last_usage = plan.get("usage") if isinstance(plan, dict) else None
                    st.session_state.step = 2
                    auth.log_action(st.session_state.username, "Analysis Started")
                    st.rerun()
                except Exception as e:
                    if logic.is_quota_error(e):
                        st.warning("⚠️ API 额度已用完，请先充值后再试。")
                    else:
                        st.error(f"处理失败: {e}")
        st.markdown("</div>", unsafe_allow_html=True)

    # ================== 步骤 2: 审核 (增加源数据透视) ==================
    elif st.session_state.step == 2:
        st.markdown(
            """<div class="w2w-card"><div class="w2w-header">📊 步骤 2: 数据核对</div><div class="w2w-desc">AI 已从源文件中提取数据。</div>""",
            unsafe_allow_html=True)

        # 新增：查看 AI 读到了什么
        with st.expander("🔍 [调试] 查看 AI 读取到的源文件内容"):
            st.text_area("源文本快照", st.session_state.source_text_display, height=200, disabled=True)
            st.caption("如果这里没有你需要的数据，说明源文件格式太复杂，AI 没读出来。")

        # ======================= 【新增】核心调试功能 =======================
        # 2. JSON 结构调试窗口 (专门用来检查数据到底去哪了)
        with st.expander("🧩 [调试] 查看 AI 返回的原始 JSON (排查写入失败)"):
            st.info(
                "💡 关键检查点：\n1. 你的“社会工作/奖惩情况”是不是在 `kv` 列表里？(在 kv 才能写入大单元格)\n2. `anchor` (定位词) 的名字是不是和 Word 模板里的文字能对应上？")
            st.json(st.session_state.plan)
        # ===================================================================

        # 数据编辑器
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

        # AI 润色区
        st.markdown("""<div class="w2w-card"><div class="w2w-header">✨ AI 润色</div>""", unsafe_allow_html=True)
        c1, c2, c3 = st.columns([2, 2, 1])
        t_target = c1.selectbox("选择字段", edited_df['anchor'].tolist())
        t_prompt = c2.text_input("指令", placeholder="例如：扩充到200字，语气更自信")
        if c3.button("执行", use_container_width=True):
            try:
                client = OpenAI(api_key=api_key, base_url=model_settings["base_url"])
                idx = st.session_state.kv_df.index[st.session_state.kv_df['anchor'] == t_target].tolist()[0]
                curr = edited_df.loc[idx, 'val']
                new_val = logic.refine_text_v2(client, curr, t_prompt, model_settings=model_settings)
                st.session_state.kv_df.at[idx, 'val'] = new_val
                st.rerun()
            except Exception as e:
                if logic.is_quota_error(e):
                    st.warning("⚠️ API 额度已用完，请先充值后再试。")
                else:
                    st.error(f"润色失败: {e}")
        st.markdown("</div>", unsafe_allow_html=True)

        c_b1, c_b2 = st.columns(2)
        if c_b1.button("🔙 返回重传"):
            st.session_state.step = 1
            st.rerun()
        if c_b2.button("✅ 确认生成", type="primary"):
            st.session_state.plan['kv'] = edited_df.to_dict('records')
            st.session_state.preview_bytes = None
            st.session_state.preview_data = None
            st.session_state.step = 3
            st.rerun()

    # ================== 步骤 3: 预览 (生成前检查) ==================
    elif st.session_state.step == 3:
        st.markdown(
            """<div class="w2w-card"><div class="w2w-header">👀 步骤 3: 预览确认</div><div class="w2w-desc">在下载前先查看预览效果，确认无误再生成最终结果。</div>""",
            unsafe_allow_html=True)

        bar = st.progress(0)
        try:
            p_template = os.path.join("temp", "target_template.docx")
            p_out = os.path.join("temp", "final_result.docx")

            # 强制恢复文件
            if st.session_state.get('template_bytes'):
                if not os.path.exists("temp"): os.makedirs("temp")
                with open(p_template, "wb") as f:
                    f.write(st.session_state.template_bytes)
            else:
                st.error("⚠️ 会话过期")
                if st.button("🔙 返回首页"):
                    st.session_state.step = 1
                    st.rerun()
                st.stop()

            if st.session_state.preview_bytes is None:
                def update_bar(p, msg):
                    bar.progress(p, text=msg)
                    time.sleep(0.05)

                logic.execute_word_writing_v2(st.session_state.plan, p_template, p_out, progress_callback=update_bar)
                with open(p_out, "rb") as f:
                    st.session_state.preview_bytes = f.read()
                st.session_state.preview_data = logic.get_docx_preview_data(p_out)

            paragraphs, tables = st.session_state.preview_data or ([], [])
            with st.expander("📄 文本预览 (格式仅供参考)", expanded=True):
                if paragraphs:
                    st.text_area("预览内容", "\n".join(paragraphs), height=260, disabled=True)
                else:
                    st.info("未检测到正文文本预览。")

            if tables:
                for idx, table in enumerate(tables, start=1):
                    with st.expander(f"📊 表格预览 {idx}"):
                        st.dataframe(pd.DataFrame(table), use_container_width=True)
            else:
                st.info("未检测到表格预览。")

            st.markdown("</div>", unsafe_allow_html=True)
            col_back, col_confirm, col_new = st.columns([2, 2, 2])
            if col_back.button("✏️ 返回修改"):
                st.session_state.step = 2
                st.rerun()
            if col_confirm.button("✅ 确认生成"):
                st.session_state.step = 4
                st.rerun()
            if col_new.button("🔄 开始新任务"):
                st.session_state.step = 1
                st.session_state.plan = None
                st.session_state.preview_bytes = None
                st.session_state.preview_data = None
                if 'auto_profile_name' in st.session_state:
                    del st.session_state.auto_profile_name
                st.rerun()

        except Exception as e:
            st.error(f"预览生成失败: {e}")
            st.markdown("### ⚠️ 遇到问题了？")
            if st.button("🔙 返回第一步 (重新上传)", type="primary"):
                st.session_state.step = 1
                st.rerun()

    # ================== 步骤 4: 写入下载 (增加错误回退) ==================
    elif st.session_state.step == 4:
        st.markdown(
            """<div class="w2w-card" style="text-align:center; padding:40px;"><h3 style="color:#4F46E5;">✅ 预览完成，可下载结果</h3></div>""",
            unsafe_allow_html=True)
        try:
            if not st.session_state.preview_bytes:
                st.error("⚠️ 预览数据不存在，请返回上一步重新生成。")
                if st.button("🔙 返回预览"):
                    st.session_state.step = 3
                    st.rerun()
                st.stop()

            output_name = f"WordToWord_V1.0_{st.session_state.user_filename_display}"
            col_dl, col_back, col_new = st.columns([3, 2, 2])
            col_dl.download_button(
                "📥 下载结果",
                st.session_state.preview_bytes,
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True,
            )

            if col_back.button("✏️ 不满意？返回修改"):
                st.session_state.step = 2
                st.rerun()

            if col_new.button("🔄 开始新任务"):
                st.session_state.step = 1
                if 'auto_profile_name' in st.session_state:
                    del st.session_state.auto_profile_name
                st.session_state.plan = None
                st.session_state.preview_bytes = None
                st.session_state.preview_data = None
                st.rerun()

        except Exception as e:
            st.error(f"写入出错: {e}")
            st.markdown("### ⚠️ 遇到问题了？")
            if st.button("🔙 返回第一步 (重新上传)", type="primary"):
                st.session_state.step = 1
                st.rerun()


# ================= 路由 =================
if not st.session_state.logged_in:
    login_page()
else:
    if st.session_state.user_role == 'admin':
        admin_page()
    else:
        user_page()
