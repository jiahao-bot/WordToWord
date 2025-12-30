import json
import re
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from copy import deepcopy
from openai import OpenAI
import os
import zipfile  # <--- 【新增】引入这个库用来检测文件坏没坏

# 尝试导入 pdfplumber
try:
    import pdfplumber
except ImportError:
    pdfplumber = None


# ================= 文本读取逻辑 =================

def _read_pdf(file_path):
    """专门读取 PDF 文本"""
    if pdfplumber is None:
        return "[系统错误] 缺少 pdfplumber 库，请运行 pip install pdfplumber 安装。"

    text_content = []
    try:
        with pdfplumber.open(file_path) as pdf:
            for i, page in enumerate(pdf.pages):
                txt = page.extract_text()
                if txt:
                    text_content.append(f"[PDF_第{i + 1}页] {txt}")
                tables = page.extract_tables()
                for t_idx, table in enumerate(tables):
                    table_str = " | ".join([" ".join([str(c) for c in row if c]) for row in table])
                    if table_str:
                        text_content.append(f"[PDF_表格_{i + 1}_{t_idx}] {table_str}")
    except Exception as e:
        return f"[PDF读取失败] {str(e)}"

    return "\n".join(text_content)


def read_file_content(file_path):
    """通用文件读取"""
    ext = os.path.splitext(file_path)[1].lower()

    if ext == '.pdf':
        return _read_pdf(file_path)

    # 增加 docx 格式预检
    if not zipfile.is_zipfile(file_path):
        # 如果不是 PDF 也不是合法的 zip (docx本质是zip)，尝试作为纯文本或报错
        return f"[严重警告] 文件 '{os.path.basename(file_path)}' 不是有效的 .docx 格式。\n请不要直接修改后缀名，请用 Word 打开后‘另存为’ .docx。"

    try:
        doc = Document(file_path)
        text = []
        for i, table in enumerate(doc.tables):
            for row in table.rows:
                row_txt = " | ".join([c.text.strip() for c in row.cells if c.text.strip()])
                if row_txt: text.append(f"[表格_{i}] {row_txt}")

        for p in doc.paragraphs:
            if p.text.strip(): text.append(f"[段落] {p.text.strip()}")

        return "\n".join(text)
    except Exception as e:
        return f"[文件读取错误] {str(e)}"


# ================= V2 核心 Prompt =================
def generate_filling_plan_v2(client, old_data, target_structure):
    prompt = f"""
    你是一个高级数据提取与合成引擎 (V1.0)。

    【源数据】
    {old_data[:8000]} 

    【目标结构】
    {target_structure[:3000]}

    【指令】
    1. **KV提取**: 提取事实性数据。
    2. **智能合成**: 遇到"自我评价"等主观项，若源数据无对应，请根据经历**自动创作**一段。
    3. **列表全量提取**: 课程、奖项、论文，必须全部提取。
    4. **勾选框**: 识别“有/无”状态。

    【输出 JSON】
    {{
        "kv": [
            {{"anchor": "姓名", "val": "张三", "source": "EXTRACTED"}}
        ],
        "checkbox": [{{"keyword": "不及格", "status": "无"}}],
        "lists": [
             {{ "keyword": "课程类型", "headers": ["名称", "成绩"], "data": [["数学", "90"]] }}
        ]
    }}
    """
    response = client.chat.completions.create(
        model="deepseek-chat",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.7
    )
    content = response.choices[0].message.content
    content = re.sub(r'```json\s*|\s*```', '', content)
    try:
        return json.loads(content)
    except:
        # 简单的容错，防止 JSON 解析失败导致崩溃
        return {"kv": [], "checkbox": [], "lists": []}


def refine_text_v2(client, original_text, instruction):
    prompt = f"请根据用户指令润色以下文本。\n【原文】{original_text}\n【指令】{instruction}\n【输出】仅输出润色后的文本。"
    response = client.chat.completions.create(
        model="deepseek-chat",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.7
    )
    return response.choices[0].message.content


# ================= V2 写入逻辑 =================

def force_write_cell(cell, text):
    original_size = None
    if cell.paragraphs and cell.paragraphs[0].runs:
        original_size = cell.paragraphs[0].runs[0].font.size
    cell.text = ""
    p = cell.paragraphs[0]
    run = p.add_run(str(text))
    run.font.color.rgb = RGBColor(0, 0, 0)
    run.font.name = '宋体'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    if original_size:
        run.font.size = original_size
    else:
        run.font.size = Pt(10.5)


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
    if status in ["无", "No"]:
        if "无□" in text:
            new_text = text.replace("无□", "无☑")
        elif "无" in text and "□" in text:
            new_text = text.replace("无□", "无☑").replace("无 □", "无 ☑")
    elif status in ["有", "Yes"]:
        if "有□" in text:
            new_text = text.replace("有□", "有☑")
        elif "有" in text and "□" in text:
            new_text = text.replace("有□", "有☑").replace("有 □", "有 ☑")
    if new_text != text:
        force_write_cell(cell, new_text)
        return True
    return False


def deepcopy_row(table, source_row):
    tbl = table._tbl
    tr = deepcopy(source_row._tr)
    source_row._tr.addprevious(tr)
    return table.rows[source_row._index - 1]


def execute_word_writing_v2(plan, template_path, output_path, progress_callback=None):
    # 【新增】这里就是核心修复！在打开文件前，先检查它是不是真正的 Docx
    if not zipfile.is_zipfile(template_path):
        raise ValueError(
            "❌ 上传的【目标文件】格式错误！\n它可能是一个旧版 .doc 文件被直接改了后缀名。\n\n💡 解决方法：\n请在电脑上用 Word 打开该文件，点击‘文件’ -> ‘另存为’ -> 选择 ‘Word 文档 (*.docx)’，然后上传新保存的文件。")

    doc = Document(template_path)

    # 1. KV
    if progress_callback: progress_callback(10, "正在写入基础信息...")
    for item in plan.get("kv", []):
        anchor, val = item["anchor"], item["val"]
        found = False
        for table in doc.tables:
            for row in table.rows:
                for c_idx, cell in enumerate(row.cells):
                    if anchor in cell.text.replace(" ", ""):
                        target = get_next_distinct_cell(row, c_idx)
                        if target and anchor not in target.text:
                            force_write_cell(target, val)
                            found = True;
                            break
                if found: break
            if found: break

    # 2. Checkbox
    if progress_callback: progress_callback(40, "正在处理勾选框...")
    for item in plan.get("checkbox", []):
        keyword, status = item["keyword"], item["status"]
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if keyword in cell.text:
                        if "□" in cell.text:
                            handle_checkbox(cell, status)
                        else:
                            for c in row.cells:
                                if "□" in c.text: handle_checkbox(c, status)

    # 3. Lists
    if progress_callback: progress_callback(70, "正在克隆列表数据...")
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
                    found_template = False
                    for offset in range(1, 4):
                        if r_idx + offset < len(table.rows):
                            check_row_txt = "".join([c.text for c in table.rows[r_idx + offset].cells])
                            if "写全称" in check_row_txt or "xx" in check_row_txt.lower():
                                target_table = table
                                template_row_idx = r_idx + offset
                                found_template = True
                                break
                    if not found_template and r_idx + 1 < len(table.rows):
                        target_table = table
                        template_row_idx = r_idx + 1
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
                        force_write_cell(distinct_cells[i], val)
            target_table._tbl.remove(template_row._tr)

    doc.save(output_path)
    if progress_callback: progress_callback(100, "完成")