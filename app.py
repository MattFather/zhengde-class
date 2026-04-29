import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io
import datetime

# ================= 網頁整體設定 =================
st.set_page_config(page_title="正德國中 - 調/代 課單系統", layout="wide")
st.title("🏫 正德國中 - 調/代 課單自動對調系統")

# ================= 核心輔助函式 =================
def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    for side in ["top", "left", "bottom", "right"]:
        if side in kwargs:
            tag = 'w:{}'.format(side)
            element = tcPr.find(qn(tag))
            if element is not None:
                tcPr.remove(element)
            element = OxmlElement(tag)
            for key, val in kwargs[side].items():
                element.set(qn('w:{}'.format(key)), str(val))
            tcPr.append(element)

def set_chinese_font(doc, font_name='標楷體'):
    doc.styles['Normal'].font.name = font_name
    doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

def generate_timetable_block(container_cell, title_suffix, sch_year, sch_term, class_label, filtered_df, is_teacher_side=True):
    # 1. 標題 (14pt)
    p_header = container_cell.paragraphs[0]
    p_header.paragraph_format.space_before = Pt(0)
    p_header.paragraph_format.space_after = Pt(0)
    p_header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_h = p_header.add_run(f"新北市立正德國民中學 {sch_year}學年度第{sch_term}學期\n調/代 課單")
    run_h.bold = True
    run_h.font.size = Pt(14) 

    # 2. 班級 (12pt)
    p_class = container_cell.add_paragraph()
    p_class.paragraph_format.space_before = Pt(0)
    p_class.paragraph_format.space_after = Pt(0)
    p_class.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_class = p_class.add_run(f"班級：{class_label}")
    run_class.bold = True
    run_class.font.size = Pt(12) 

    # 3. 建立表格
    inner_table = container_cell.add_table(rows=9, cols=6)
    inner_table.style = 'Table Grid'
    
    weekdays_list = ["一", "二", "三", "四", "五"]
    h_cells = inner_table.rows[0].cells
    
    # 左上角 (節次/星期) 11pt
    p_tl = h_cells[0].paragraphs[0]
    p_tl.text = ""
    run_tl = p_tl.add_run("節次/星期")
    run_tl.font.size = Pt(11)
    
    # 星期 12pt
    for i, day in enumerate(weekdays_list):
        p_day = h_cells[i+1].paragraphs[0]
        p_day.text = ""
        run_day = p_day.add_run(day)
        run_day.font.size = Pt(12)

    # 4. 節次(12pt) 與 時間(9pt 橫式)
    periods_list = ["1", "2", "3", "4", "5", "6", "7", "8"]
    times_list = ["08:20-09:05", "09:15-10:00", "10:10-10:55", "11:05-11:50", "13:00-13:45", "13:55-14:40", "15:00-15:45", "15:55-16:40"]
    for r_idx in range(8):
        row = inner_table.rows[r_idx + 1]
        row.height = Cm(1.5)
        cell_p = row.cells[0].paragraphs[0]
        cell_p.paragraph_format.space_after = Pt(0)
        run_num = cell_p.add_run(periods_list[r_idx])
        run_num.bold = True
        run_num.font.size = Pt(12) 
        cell_p.add_run("\n")
        run_time = cell_p.add_run(times_list[r_idx])
        run_time.font.size = Pt(9) 

    # 5. 填入課程資料
    day_map = {0: 1, 1: 2, 2: 3, 3: 4, 4: 5} 
    for _, row_data in filtered_df.iterrows():
        if pd.notnull(row_data["日期"]):
            d_idx = day_map.get(row_data["日期"].weekday())
            try:
                p_idx = int(str(row_data["節次"]).split()[1])
            except: continue
            if d_idx and p_idx:
                cell = inner_table.rows[p_idx].cells[d_idx]
                cell.text = "" 
                
                c_name = str(row_data["班級"]).strip() if pd.notnull(row_data["班級"]) else ""
                s_name = str(row_data["科目"]).strip() if pd.notnull(row_data["科目"]) else ""
                
                p1 = cell.paragraphs[0]
                p1.paragraph_format.space_after = Pt(0)
                p1.add_run(row_data["日期"].strftime("%m/%d"))
                
                p2 = cell.add_paragraph()
                p2.paragraph_format.space_after = Pt(0)
                subj_display = f"{c_name} {s_name}".strip() if is_teacher_side and c_name else s_name
                p2.add_run(subj_display)
                
                p3 = cell.add_paragraph()
                p3.paragraph_format.space_after = Pt(0)
                p3.add_run(str(row_data["老師"]))
                
                p4 = cell.add_paragraph()
                p4.paragraph_format.space_after = Pt(0)
                
                if row_data["調/代課"] == "代課":
                    run_type = p4.add_run("[代課]")
                    run_type.font.size = Pt(9)
                else:
                    if is_teacher_side and pd.notnull(row_data.get("原資訊")):
                        run_type = p4.add_run(str(row_data["原資訊"]))
                        run_type.font.size = Pt(8) 
                    else:
                        run_type = p4.add_run("[調課]")
                        run_type.font.size = Pt(9)

    # 6. 表格格式
    for r in range(9):
        for c in range(6):
            curr_cell = inner_table.rows[r].cells[c]
            curr_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            for para in curr_cell.paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                para.paragraph_format.line_spacing = 1.0
            if r == 4: set_cell_border(curr_cell, bottom={"sz": 24, "val": "single", "color": "#000000"})
            if r == 5: set_cell_border(curr_cell, top={"sz": 24, "val": "single", "color": "#000000"})

    # 7. 列印日期 (10pt, 物理定位消除空白)
    print_p = container_cell.paragraphs[-1] 
    print_p.text = ""
    print_p.paragraph_format.space_before = Pt(0) 
    print_p.paragraph_format.space_after = Pt(0)
    print_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run_date = print_p.add_run(f"列印：{datetime.date.today().strftime('%Y/%m/%d')}")
    run_date.font.size = Pt(10)

    # 8. 腳註
    footer_p = container_cell.add_paragraph()
    footer_p.paragraph_format.space_before = Pt(0)
    footer_p.paragraph_format.space_after = Pt(0)
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_f = footer_p.add_run(f"({title_suffix})")
    run_f.bold = True
    run_f.font.size = Pt(10)

# ================= 自動對調處理引擎 =================
def process_swap_logic(df):
    df_result = []
    w_map = {0:"一", 1:"二", 2:"三", 3:"四", 4:"五", 5:"六", 6:"日"}
    
    # 1. 處理代課
    subs = df[df["調/代課"] == "代課"].copy()
    for _, r in subs.iterrows():
        df_result.append(r)
        
    # 2. 處理調課
    swaps = df[df["調/代課"] == "調課"].copy()
    pair_ids = [p for p in swaps["配對編號"].unique() if pd.notnull(p) and str(p).strip() != ""]
    
    for pid in pair_ids:
        rows = swaps[swaps["配對編號"] == pid].sort_index()
        n = len(rows)
        if n >= 2:
            orig_dates = rows["日期"].tolist()
            orig_periods = rows["節次"].tolist()
            
            shifted_dates = orig_dates[1:] + [orig_dates[0]]
            shifted_periods = orig_periods[1:] + [orig_periods[0]]
            
            for i in range(n):
                new_row = rows.iloc[i].copy()
                o_d = orig_dates[i]
                o_p = orig_periods[i]
                w_day = w_map.get(o_d.weekday(), "")
                new_row["原資訊"] = f"[原{o_d.strftime('%m/%d')}({w_day}){o_p}]"
                
                new_row["日期"] = shifted_dates[i]
                new_row["節次"] = shifted_periods[i]
                df_result.append(new_row)
        else:
            for _, r in rows.iterrows():
                df_result.append(r)
    
    # 3. 處理未填編號的調課
    no_id = swaps[swaps["配對編號"].isna() | (swaps["配對編號"] == "")]
    for _, r in no_id.iterrows():
        df_result.append(r)
        
    return pd.DataFrame(df_result)

def create_docx(sch_year, sch_term, edited_df):
    doc = Document()
    section = doc.sections[0]
    section.orient = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width
    section.left_margin = section.right_margin = section.top_margin = section.bottom_margin = Cm(0.5)
    set_chinese_font(doc, '標楷體')

    df_raw = edited_df[edited_df["勾選列印資料"] == True].copy()
    if df_raw.empty: return None
    
    df_raw["配對編號"] = df_raw["配對編號"].fillna("").astype(str).str.strip()
    df_raw["班級"] = df_raw["班級"].fillna("").astype(str).str.strip()
    df_raw["老師"] = df_raw["老師"].fillna("").astype(str).str.strip()
    
    df_processed = process_swap_logic(df_raw)

    all_blocks = []
    # A. 存查聯
    classes = sorted(list(set([c for c in df_processed["班級"] if c != ""])))
    all_blocks.append({"suffix": "存查聯", "label": ", ".join(classes), "df": df_processed, "is_teacher": True})

    # B. 通知聯
    teachers = sorted(list(set([t for t in df_processed["老師"] if t != ""])))
    for t in teachers:
        df_t = df_processed[df_processed["老師"] == t]
        t_classes = sorted(list(set([c for c in df_t["班級"] if c != ""])))
        all_blocks.append({"suffix": f"通知聯 - {t} 老師", "label": ", ".join(t_classes), "df": df_t, "is_teacher": True})

    # C. 公告聯
    for c in classes:
        df_c = df_processed[df_processed["班級"] == c]
        all_blocks.append({"suffix": "公告聯", "label": c, "df": df_c, "is_teacher": False})

    for i in range(0, len(all_blocks), 2):
        if i > 0: doc.add_page_break()
        table = doc.add_table(rows=1, cols=2)
        table.width = Cm(28.7)
        b1 = all_blocks[i]
        generate_timetable_block(table.cell(0, 0), b1["suffix"], sch_year, sch_term, b1["label"], b1["df"], is_teacher_side=b1["is_teacher"])
        if i + 1 < len(all_blocks):
            b2 = all_blocks[i+1]
            generate_timetable_block(table.cell(0, 1), b2["suffix"], sch_year, sch_term, b2["label"], b2["df"], is_teacher_side=b2["is_teacher"])
        else:
            p = table.cell(0, 1).paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run("\n\n\n\n\n\n(裁切線)\n----------\n正德教務處專用")

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# ================= 網頁介面 =================
st.markdown("### 📅 調/代 課單自動對調系統 (完整版)")
c1, c2 = st.columns(2)
with c1: sch_year = st.text_input("學年度", value="114")
with c2: sch_term = st.selectbox("學期", ["一", "二"], index=1)

# ★ 修正後的精準說明區塊 ★
st.info("""
💡 **操作說明**：（點選表格「**左側**」方塊後按 `Delete` 鍵可刪除不需要的資料列）
1. **代課**：類型選 `[代課]`，`[配對編號]` **留空**。系統會保留原上課時間，僅更換老師與科目。
2. **互調**：類型選 `[調課]`，將要互調的兩筆原始資料 `[配對編號]` 填入 **相同數字**（如：1）。系統會自動互換兩者的時間，並在通知單上貼心標註原上課時間。
3. **多角調**：類型選 `[調課]`，將涉及的所有原始資料 `[配對編號]` 填入 **相同數字**。系統會依輸入順序自動循環對調（第1列 ➔ 第2列位置，第2列 ➔ 第3列位置...最後一列 ➔ 第1列位置）。
""")

if 'res_data' not in st.session_state:
    st.session_state.res_data = pd.DataFrame([
        {"勾選列印資料": True, "配對編號": "1", "班級": "717", "日期": datetime.date(2026, 5, 11), "節次": "第 3 節", "科目": "生物", "老師": "生物老師", "調/代課": "調課"},
        {"勾選列印資料": True, "配對編號": "1", "班級": "717", "日期": datetime.date(2026, 5, 15), "節次": "第 6 節", "科目": "數學", "老師": "數學老師", "調/代課": "調課"}
    ])

edited_df = st.data_editor(
    st.session_state.res_data,
    column_config={
        "勾選列印資料": st.column_config.CheckboxColumn("勾選"),
        "配對編號": st.column_config.TextColumn("配對編號"),
        "班級": st.column_config.TextColumn("班級"),
        "日期": st.column_config.DateColumn("日期", format="MM/DD"),
        "節次": st.column_config.SelectboxColumn("節次", options=[f"第 {i} 節" for i in range(1, 9)]),
        "科目": st.column_config.TextColumn("科目"),
        "老師": st.column_config.TextColumn("老師"),
        "調/代課": st.column_config.SelectboxColumn("調/代課", options=["調課", "代課"]),
    },
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,
    column_order=("勾選列印資料", "配對編號", "班級", "日期", "節次", "科目", "老師", "調/代課")
)

st.divider()
data = create_docx(sch_year, sch_term, edited_df)
if data:
    st.download_button(
        label="📥 下載【完整優化版】調代課單 (Word)",
        data=data,
        file_name=f"正德調代課單_完整版_{datetime.date.today().strftime('%Y%m%d')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
        use_container_width=True
    )