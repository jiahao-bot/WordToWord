import json
import re
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from copy import deepcopy
from openai import OpenAI
import os
import zipfile
import difflib

try:
    import pdfplumber
except ImportError:
    pdfplumber = None


# ================== 文件格式预检 (保持不变) ==================
def validate_file_format(file_path):
    if not os.path.exists(file_path):
        return False, "文件上传失败，未能保存到服务器。"
    filename = os.path.basename(file_path)
    ext = os.path.splitext(filename)[1].lower()

    if ext == '.docx':
        if not zipfile.is_zipfile(file_path):
            return False, f"❌ 文件【{filename}】格式错误！\n它看起来像是旧版 .doc 或已损坏。\n💡 请用 Word 打开并‘另存为’ .docx 格式。"
        try:
            Document(file_path)
        except Exception as e:
            return False, f"❌ 文件【{filename}】内容损坏: {str(e)}"
    elif ext == '.pdf':
        if pdfplumber is None:
            return False, "缺少 pdfplumber 库。"
        try:
            with pdfplumber.open(file_path) as pdf:
                if len(pdf.pages) == 0: return False, "❌ PDF 文件是空的。"
        except Exception as e:
            return False, f"❌ PDF 损坏: {str(e)}"
    return True, "OK"


# ================= 文本读取 (保持不变) =================
def _read_pdf(file_path):
    if pdfplumber is None: return ""
    text_content = []
    try:
        with pdfplumber.open(file_path) as pdf:
            for i, page in enumerate(pdf.pages):
                txt = page.extract_text()
                if txt: text_content.append(f"[PDF_第{i + 1}页] {txt}")
                tables = page.extract_tables()
                for t_idx, table in enumerate(tables):
                    clean_table = []
                    for row in table:
                        clean_row = [str(c).replace('\n', ' ') for c in row if c]
                        if clean_row: clean_table.append(" | ".join(clean_row))
                    if clean_table:
                        text_content.append(f"[PDF_表格_{i + 1}_{t_idx}]\n" + "\n".join(clean_table))
    except Exception as e:
        return f"[PDF读取失败] {str(e)}"
    return "\n".join(text_content)


def read_file_content(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    if ext == '.pdf': return _read_pdf(file_path)
    try:
        doc = Document(file_path)
        text = []
        for i, table in enumerate(doc.tables):
            table_data = []
            for row in table.rows:
                row_txt = " | ".join([c.text.strip() for c in row.cells if c.text.strip()])
                if row_txt: table_data.append(row_txt)
            if table_data:
                text.append(f"【表格区_{i}】\n" + "\n".join(table_data))

        para_data = []
        for p in doc.paragraphs:
            if p.text.strip(): para_data.append(p.text.strip())
        if para_data:
            text.append("【正文区】\n" + "\n".join(para_data))

        return "\n\n".join(text)
    except Exception as e:
        return f"[读取错误] {str(e)}"


# ================= V5 核心 Prompt (修复基础信息遗漏) =================
def generate_filling_plan_v2(client, old_data, target_structure):
    prompt = f"""
    你是一个专业的数据迁移专家。

    【源数据】
    {old_data[:12000]} 

    【目标表结构】
    {target_structure[:4000]}

    【必须严格执行的指令】
    1. **全面提取 KV (基础信息 + 软信息)**:
       - **基础信息**: 必须地毯式提取所有短字段！包括“学号”、“性别”、“民族”、“籍贯”、“政治面貌”、“出生年月”等。不要因为它们简单就忽略！
       - **软信息**: 对于“自我鉴定”、“主要事迹”等长文本，如果源数据没有，请**根据简历事实自动撰写**，禁止留空。

    2. **Lists (多行表格)**:
       - 凡是目标表中有明确表头（如：时间|课程|成绩）的，必须提取为 `lists`。
       - **严格对齐**: `headers` 列数必须与 `data` 列数一致。

    3. **Checkbox (勾选框)**:
       - 寻找“□”符号。
       - 输出 keyword (选项文字) 和 status (有/无/是/否)。

    【输出格式 (JSON)】
    {{
        "kv": [
            {{"anchor": "姓名", "val": "张三"}},
            {{"anchor": "学号", "val": "20201101"}},
            {{"anchor": "性别", "val": "男"}},
            {{"anchor": "自我鉴定", "val": "本人在校期间..."}}
        ],
        "checkbox": [
            {{"keyword": "党员", "status": "有"}},
            {{"keyword": "英语六级", "status": "无"}}
        ],
        "lists": [
            {{
                "keyword": "获奖情况", 
                "headers": ["时间", "奖项", "等级"],
                "data": [["2023.09", "一等奖", "校级"]]
            }}
        ]
    }}
    """
    response = client.chat.completions.create(
        model="deepseek-chat",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.25  # 微调温度，平衡创造性(软信息)和准确性(基础信息)
    )
    content = response.choices[0].message.content
    content = re.sub(r'```json\s*|\s*```', '', content)

    try:
        plan = json.loads(content)

        # 自动清洗表格列数 (防止报错)
        if "lists" in plan:
            for lst in plan["lists"]:
                headers = lst.get("headers", [])
                data = lst.get("data", [])
                if headers and data:
                    num_cols = len(headers)
                    cleaned_data = []
                    for row in data:
                        if len(row) > num_cols:
                            cleaned_data.append(row[:num_cols])
                        elif len(row) < num_cols:
                            cleaned_data.append(row + [""] * (num_cols - len(row)))
                        else:
                            cleaned_data.append(row)
                    lst["data"] = cleaned_data

        return plan
    except:
        return {"kv": [], "checkbox": [], "lists": []}


def refine_text_v2(client, original_text, instruction):
    prompt = f"原文：{original_text}\n指令：{instruction}\n请输出修改后的结果："
    response = client.chat.completions.create(
        model="deepseek-chat", messages=[{"role": "user", "content": prompt}]
    )
    return response.choices[0].message.content


# ================= 写入逻辑 =================

def force_write_cell(cell, text, alignment="auto"):
    """
    格式美化：清除原有格式，自动判断居中或左对齐
    """
    # 保护性检查：如果 text 是 None，转为空字符串
    if text is None: text = ""

    cell._element.clear_content()
    p = cell.add_paragraph()

    text_len = len(str(text))
    if alignment == "auto":
        # 短文本居中，长文本左对齐
        if text_len < 15 and "\n" not in str(text):
            p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        else:
            p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    run = p.add_run(str(text))
    run.font.name = '宋体'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    run.font.size = Pt(10.5)
    run.font.color.rgb = RGBColor(0, 0, 0)


def get_next_distinct_cell(row, current_idx):
    current_cell = row.cells[current_idx]
    for i in range(current_idx + 1, len(row.cells)):
        next_cell = row.cells[i]
        if next_cell._element is not current_cell._element:
            return next_cell
    return None


def handle_checkbox(cell, status):
    text = cell.text
    new_text = text

    # ====== 【关键修复】: 使用穷举匹配法，解决“有□ 无□”同格问题 ======

    # 1. 映射表：状态 -> 需要替换的目标字符串
    replace_map = {}

    if status in ["无", "No", "否", "None", "未通过"]:
        replace_map = {
            "无□": "无☑", "无 □": "无 ☑",
            "否□": "否☑", "否 □": "否 ☑",
            "未通过□": "未通过☑"
        }
    elif status in ["有", "Yes", "是", "Have", "通过"]:
        replace_map = {
            "有□": "有☑", "有 □": "有 ☑",
            "是□": "是☑", "是 □": "是 ☑",
            "通过□": "通过☑"
        }

    # 2. 优先尝试包含文字的精确替换 (解决 "有□ 无□" 这种场景)
    replaced_flag = False
    for k, v in replace_map.items():
        if k in text:
            new_text = new_text.replace(k, v)
            replaced_flag = True

    # 3. 如果没找到带文字的框，但确实是勾选状态，且格子里只有一个框，则兜底替换
    if not replaced_flag and "□" in text:
        # 只有当状态明确为肯定，或者明确针对该项时才打钩
        if status in ["有", "Yes", "是", "Have", "True"]:
            new_text = text.replace("□", "☑", 1)  # 只替第一个，防止误伤

    if new_text != text:
        cell.text = new_text
        return True
    return False


def deepcopy_row(table, source_row):
    tbl = table._tbl
    tr = deepcopy(source_row._tr)
    source_row._tr.addprevious(tr)
    return table.rows[source_row._index - 1]


def get_fuzzy_score(anchor, target_text):
    a = anchor.replace(" ", "").replace("\n", "").lower()
    t = target_text.replace(" ", "").replace("\n", "").lower()
    if not a or not t: return 0.0
    if a == t: return 1.0
    if a in t: return 1.0
    return difflib.SequenceMatcher(None, a, t).ratio()


def execute_word_writing_v2(plan, template_path, output_path, progress_callback=None):
    if not zipfile.is_zipfile(template_path):
        raise ValueError("目标文件格式错误")
    doc = Document(template_path)

    # ---------------- 1. KV 写入 ----------------
    total_kv = len(plan.get("kv", []))
    for i, item in enumerate(plan.get("kv", [])):
        anchor, val = item["anchor"], item["val"]
        if not val: continue

        if progress_callback: progress_callback(int(10 + (i / total_kv) * 30), f"正在写入: {anchor}...")

        found = False
        for table in doc.tables:
            for row in table.rows:
                for c_idx, cell in enumerate(row.cells):
                    cell_text = cell.text.strip().replace(" ", "")
                    clean_anchor = anchor.strip().replace(" ", "")
                    match_score = get_fuzzy_score(clean_anchor, cell_text)

                    if match_score > 0.8:
                        target_cell = None
                        # 大格子逻辑 (自我鉴定)
                        if len(cell_text) > 20 or "此栏" in cell_text or "填写" in cell_text:
                            target_cell = cell
                            # 普通 KV 逻辑 (学号、姓名)
                        else:
                            candidate = get_next_distinct_cell(row, c_idx)
                            if candidate: target_cell = candidate

                        if target_cell:
                            # 保护机制：防止覆盖表头
                            # 如果目标格子很短，且包含冒号或看起来像另一个表头，跳过
                            if len(target_cell.text) < 10 and ("：" in target_cell.text or ":" in target_cell.text):
                                pass
                            else:
                                force_write_cell(target_cell, val, alignment="auto")
                                found = True
                                break
                if found: break
            if found: break

    # ---------------- 2. Checkbox 写入 (新版匹配逻辑) ----------------
    if progress_callback: progress_callback(60, "处理勾选框...")
    for item in plan.get("checkbox", []):
        keyword, status = item["keyword"], item["status"]
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    # 只有当关键字匹配时才尝试打钩
                    if keyword in cell.text:
                        handle_checkbox(cell, status)

    # ---------------- 3. Lists 写入 (保持稳定) ----------------
    if progress_callback: progress_callback(80, "处理表格列表...")
    for item in plan.get("lists", []):
        keyword = item["keyword"]
        data = item["data"]
        if not data: continue

        target_table = None
        template_row_idx = -1

        for table in doc.tables:
            for r_idx, row in enumerate(table.rows):
                row_txt = "".join([c.text for c in row.cells])
                if keyword in row_txt:
                    for offset in range(1, 4):
                        if r_idx + offset < len(table.rows):
                            check_row = table.rows[r_idx + offset]
                            check_txt = "".join([c.text for c in check_row.cells])
                            if "xx" in check_txt.lower() or len(check_txt.strip()) < 5:
                                target_table = table
                                template_row_idx = r_idx + offset
                                break
                    if target_table: break
            if target_table: break

        if target_table:
            template_row = target_table.rows[template_row_idx]

            for data_row in data:
                new_row = deepcopy_row(target_table, template_row)
                distinct_cells = []
                if len(new_row.cells) > 0:
                    distinct_cells.append(new_row.cells[0])
                    for i in range(1, len(new_row.cells)):
                        if new_row.cells[i]._element is not new_row.cells[i - 1]._element:
                            distinct_cells.append(new_row.cells[i])

                for i, val in enumerate(data_row):
                    if i < len(distinct_cells):
                        force_write_cell(distinct_cells[i], val, alignment="auto")

            target_table._tbl.remove(template_row._tr)

    doc.save(output_path)
    if progress_callback: progress_callback(100, "完成")