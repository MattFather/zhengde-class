import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Cm, Pt
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io
import datetime
import subprocess
import tempfile
import os

# ================= 雲端 PDF 轉換引擎 =================
def docx_to_pdf(docx_bytes):
    """在 Linux 伺服器背後呼叫 LibreOffice 進行轉換"""
    with tempfile.TemporaryDirectory() as tmpdir:
        docx_path = os.path.join(tmpdir, "temp.docx")
        pdf_path = os.path.join(tmpdir, "temp.pdf")
        
        # 將記憶體中的 Word 寫入暫存硬碟
        with open(docx_path, "wb") as f:
            f.write(docx_bytes)
            
        # 下達 Linux 終端機指令
        try:
            subprocess.run([
                "libreoffice", "--headless", "--convert-to", "pdf",
                "--outdir", tmpdir, docx_path
            ], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            
            # 如果成功產生 PDF，讀取出來
            if os.path.exists(pdf_path):
                with open(pdf_path, "rb") as f:
                    return f.read()
            return None
        except Exception as e:
            st.error(f"轉換引擎發生錯誤: {e}")
            return None

# ================= 網頁整體設定 =================
st.set_page_config(page_title="正德國中 - 調/代 課單系統", layout="wide")
st.title("🏫 正德國中 - 調/代 課單系統 (V.27 絕對隱私版)")

# ================= 核心輔助函式 =================
def set_cell_border(cell, **kwargs):
    # 更嚴謹的 XML 寫入方式，確保虛線能完美生成
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.find(qn('w:tcBorders'))
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
        
    for side in ["top", "left", "bottom", "right"]:
        if side in kwargs:
            tag = 'w:{}'.format(side)
            element = tcBorders.find(qn(tag))
            if element is not None:
                tcBorders.remove(element)
            element = OxmlElement(tag)
            for key, val in kwargs[side].items():
                element.set(qn('w:{}'.format(key)), str(val))
            tcBorders.append(element)

def set_chinese_font(doc, font_name='標楷體'):
    doc.styles['Normal'].font.name = font_name
    doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)

def generate_timetable_block(container_cell, title_suffix, sch_year, sch_term, issue_unit, class_label, filtered_df, is_teacher_side=True):
    # 1. 標題 (14pt)
    p_header = container_cell.paragraphs[0]
    p_header.paragraph_format.space_before = Pt(0)
    p_header.paragraph_format.space_after = Pt(0)
    p_header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_h = p_header.add_run(f"新北市立正德國民中學 {sch_year}學年度第{sch_term}學期\n調/代 課單")
    run_h.bold = True
    run_h.font.size = Pt(14) 

    # 2. 發放單位與班級 (12pt)
    p_sub = container_cell.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(0)
    
    # 扣除儲存格左右邊距後，內部可視安全寬度鎖定為 13.32cm
    tab_stops = p_sub.paragraph_format.tab_stops
    tab_stops.add_tab_stop(Cm(13.32), WD_TAB_ALIGNMENT.RIGHT)
    
    run_sub = p_sub.add_run(f"發放單位：{issue_unit}\t班級：{class_label}")
    run_sub.bold = True
    run_sub.font.size = Pt(12) 

    # 3. 建立主表格 (內部課表)
    inner_table = container_cell.add_table(rows=9, cols=6)
    inner_table.style = 'Table Grid'
    
    # 關閉自動排版，嚴格套用 13.32cm 總寬度
    inner_table.autofit = False 
    inner_widths = [Cm(2.22), Cm(2.22), Cm(2.22), Cm(2.22), Cm(2.22), Cm(2.22)]
    for j, width in enumerate(inner_widths):
        inner_table.columns[j].width = width
        for cell in inner_table.columns[j].cells:
            cell.width = width
    
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
                
                # 統一設定前三行字體為 9pt，並加上粗體
                p1 = cell.paragraphs[0]
                p1.paragraph_format.space_after = Pt(0)
                run_date_cell = p1.add_run(row_data["日期"].strftime("%m/%d"))
                run_date_cell.font.size = Pt(9)
                run_date_cell.bold = True
                
                p2 = cell.add_paragraph()
                p2.paragraph_format.space_after = Pt(0)
                subj_display = f"{c_name} {s_name}".strip() if is_teacher_side and c_name else s_name
                run_subj = p2.add_run(subj_display)
                run_subj.font.size = Pt(9)
                run_subj.bold = True
                
                p3 = cell.add_paragraph()
                p3.paragraph_format.space_after = Pt(0)
                run_teacher = p3.add_run(str(row_data["老師"]))
                run_teacher.font.size = Pt(9)
                run_teacher.bold = True
                
                # 第四行狀態提示改為 8pt 避免換行
                p4 = cell.add_paragraph()
                p4.paragraph_format.space_after = Pt(0)
                
                if row_data["調/代課"] == "代課":
                    run_type = p4.add_run("[代課]")
                    run_type.font.size = Pt(8)
                else:
                    if is_teacher_side and pd.notnull(row_data.get("原資訊")):
                        run_type = p4.add_run(str(row_data["原資訊"]))
                        run_type.font.size = Pt(8) 
                    else:
                        run_type = p4.add_run("[調課]")
                        run_type.font.size = Pt(8)

    # 6. 表格格式
    for r in range(9):
        for c in range(6):
            curr_cell = inner_table.rows[r].cells[c]
            curr_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            for para in curr_cell.paragraphs:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                para.paragraph_format.line_spacing = 1.0
            if r == 4: set_cell_border(curr_cell, bottom={"sz": 24, "val": "single", "color": "000000"})
            if r == 5: set_cell_border(curr_cell, top={"sz": 24, "val": "single", "color": "000000"})

    # 7. 列印日期
    print_p = container_cell.paragraphs[-1] 
    print_p.text = ""
    print_p.paragraph_format.space_before = Pt(0) 
    print_p.paragraph_format.space_after = Pt(0)
    print_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    # 列印日期的右側也精準鎖定在 13.32cm，與內部課表齊平
    tab_stops_print = print_p.paragraph_format.tab_stops
    tab_stops_print.add_tab_stop(Cm(13.32), WD_TAB_ALIGNMENT.RIGHT)
    run_date = print_p.add_run(f"\t列印：{datetime.date.today().strftime('%Y/%m/%d')}")
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
    
    subs = df[df["調/代課"] == "代課"].copy()
    for _, r in subs.iterrows():
        df_result.append(r)
        
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
    
    no_id = swaps[swaps["配對編號"].isna() | (swaps["配對編號"] == "")]
    for _, r in no_id.iterrows():
        df_result.append(r)
        
    return pd.DataFrame(df_result)

def create_docx(sch_year, sch_term, issue_unit, edited_df):
    doc = Document()
    section = doc.sections[0]
    section.orient = WD_ORIENT.LANDSCAPE
    
    # 強制設定 A4 尺寸
    section.page_width = Cm(29.7)
    section.page_height = Cm(21.0)
    
    # 【關鍵校正】：左邊界 0.8cm、右邊界 0.5cm 來抵銷實體印表機誤差
    section.left_margin = Cm(0.8)
    section.right_margin = Cm(0.5)
    section.top_margin = section.bottom_margin = Cm(0.5)
    
    set_chinese_font(doc, '標楷體')

    df_raw = edited_df[edited_df["勾選列印資料"] == True].copy()
    if df_raw.empty: return None
    
    df_raw["配對編號"] = df_raw["配對編號"].fillna("").astype(str).str.strip()
    df_raw["班級"] = df_raw["班級"].fillna("").astype(str).str.strip()
    df_raw["老師"] = df_raw["老師"].fillna("").astype(str).str.strip()
    
    df_processed = process_swap_logic(df_raw)

    all_blocks = []
    classes = sorted(list(set([c for c in df_processed["班級"] if c != ""])))
    all_blocks.append({"suffix": "存查聯", "label": ", ".join(classes), "df": df_processed, "is_teacher": True})

    teachers = sorted(list(set([t for t in df_processed["老師"] if t != ""])))
    for t in teachers:
        df_t = df_processed[df_processed["老師"] == t]
        t_classes = sorted(list(set([c for c in df_t["班級"] if c != ""])))
        all_blocks.append({"suffix": f"教師通知聯 - {t} 老師", "label": ", ".join(t_classes), "df": df_t, "is_teacher": True})

    for c in classes:
        df_c = df_processed[df_processed["班級"] == c]
        all_blocks.append({"suffix": "班級公告聯", "label": c, "df": df_c, "is_teacher": False})

    for i in range(0, len(all_blocks), 2):
        if i > 0: doc.add_page_break()
        
        # 維持 4 欄位結構：左側 13.7 + 左縫隙 0.5 + 右縫隙 0.5 + 右側 13.7 = 總寬 28.4
        table = doc.add_table(rows=1, cols=4)
        table.autofit = False
        
        col_widths = [Cm(13.7), Cm(0.5), Cm(0.5), Cm(13.7)]
        for j in range(4):
            table.columns[j].width = col_widths[j]
            for cell in table.columns[j].cells:
                cell.width = col_widths[j]

        # 在「左側縫隙 (第2欄)」的右邊緣畫上切割虛線
        set_cell_border(table.cell(0, 1), right={"sz": 6, "val": "dashed", "color": "808080"})

        b1 = all_blocks[i]
        generate_timetable_block(table.cell(0, 0), b1["suffix"], sch_year, sch_term, issue_unit, b1["label"], b1["df"], is_teacher_side=b1["is_teacher"])
        
        if i + 1 < len(all_blocks):
            b2 = all_blocks[i+1]
            generate_timetable_block(table.cell(0, 3), b2["suffix"], sch_year, sch_term, issue_unit, b2["label"], b2["df"], is_teacher_side=b2["is_teacher"])

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# ================= 網頁介面 =================
st.markdown("### 📅 調/代 課單自動對調系統 (V.27 絕對隱私版)")

# 增加發放單位輸入框
c1, c2, c3 = st.columns(3)
with c1: sch_year = st.text_input("學年度", value="114")
with c2: sch_term = st.selectbox("學期", ["一", "二"], index=1)
with c3: issue_unit = st.text_input("發放單位", value="ＯＯＯ老師")

st.info("""
💡 **操作說明**：（點選表格「**左側**」方塊後按 `Delete` 鍵可刪除不需要的資料列）
1. **代課**：類型選 `[代課]`，`[配對編號]` **留空**。保留原上課時間，僅更換老師與科目。
2. **互調**：類型選 `[調課]`，兩筆原始資料 `[配對編號]` 填入 **相同數字**。系統互換時間並標註原上課時間。
3. **多角調**：類型選 `[調課]`，涉及的資料 `[配對編號]` 填入 **相同數字**。依輸入順序自動循環對調。
""")

# ================= 初始化預設資料 =================
if 'res_data' not in st.session_state:
    st.session_state.res_data = pd.DataFrame([
        {"勾選列印資料": True, "配對編號": "1", "班級": "717", "日期": datetime.date(2026, 5, 11), "節次": "第 3 節", "科目": "生物", "老師": "王小帥", "調/代課": "調課"},
        {"勾選列印資料": True, "配對編號": "1", "班級": "717", "日期": datetime.date(2026, 5, 15), "節次": "第 6 節", "科目": "數學", "老師": "林小美", "調/代課": "調課"}
    ])

# ================= 隱私保護：本機進度存取區 =================
st.markdown("#### 🔒 本機進度存取區 (資料不留雲端，保障隱私)")
c_upload, c_empty = st.columns([1, 1])

with c_upload:
    uploaded_file = st.file_uploader("📂 如果您有之前下載的進度檔 (.csv)，請在此上傳恢復：", type=["csv"])
    if uploaded_file is not None:
        # 確保只在剛上傳時讀取一次，避免干擾後續編輯
        if 'last_uploaded_id' not in st.session_state or st.session_state.last_uploaded_id != uploaded_file.file_id:
            try:
                df_upload = pd.read_csv(uploaded_file)
                df_upload["日期"] = pd.to_datetime(df_upload["日期"], errors='coerce').dt.date
                df_upload["勾選列印資料"] = df_upload["勾選列印資料"].fillna(True).astype(bool)
                df_upload.fillna("", inplace=True)
                
                st.session_state.res_data = df_upload
                st.session_state.last_uploaded_id = uploaded_file.file_id # 記錄 ID 避免重複讀取
                st.rerun() # 重新整理畫面以顯示新資料
            except Exception as e:
                st.error(f"❌ 檔案讀取失敗: {e}")

# 預先定義好的科目清單
subject_list = [
    "", "國文", "英文", "數學", "生物", "理化", "地科", "地理", "歷史", "公民", 
    "體育", "健康", "視藝", "表藝", "音樂", "家政", "童軍", "輔導", "資訊", "生科", "本土語"
]

# 編輯表格
edited_df = st.data_editor(
    st.session_state.res_data,
    column_config={
        "勾選列印資料": st.column_config.CheckboxColumn("勾選"),
        "配對編號": st.column_config.TextColumn("配對編號"),
        "班級": st.column_config.TextColumn("班級"),
        "日期": st.column_config.DateColumn("日期", format="MM/DD"),
        "節次": st.column_config.SelectboxColumn("節次", options=[f"第 {i} 節" for i in range(1, 9)]),
        "科目": st.column_config.SelectboxColumn("科目", options=subject_list, help="雙擊格子後可下拉選擇，或直接用鍵盤打字快速搜尋"),
        "老師": st.column_config.TextColumn("老師"),
        "調/代課": st.column_config.SelectboxColumn("調/代課", options=["調課", "代課"]),
    },
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,
    column_order=("勾選列印資料", "配對編號", "班級", "日期", "節次", "科目", "老師", "調/代課")
)

# 同步更新 session_state 確保下載的是最新編輯的資料
st.session_state.res_data = edited_df

# 進度下載與清空按鈕
c_download, c_clear_btn, _ = st.columns([2, 2, 4])
with c_download:
    # 將目前的表格轉成 CSV 格式 (加上 utf-8-sig 確保 Excel 打開不會亂碼)
    csv_bytes = edited_df.to_csv(index=False).encode('utf-8-sig')
    st.download_button(
        label="💾 下載目前進度 (存至您的電腦)",
        data=csv_bytes,
        file_name=f"調代課暫存_{datetime.date.today().strftime('%Y%m%d')}.csv",
        mime="text/csv",
        use_container_width=True,
        type="primary"
    )

with c_clear_btn:
    if st.button("🗑️ 清空所有表格 (開新單)", use_container_width=True):
        empty_df = pd.DataFrame(columns=["勾選列印資料", "配對編號", "班級", "日期", "節次", "科目", "老師", "調/代課"])
        st.session_state.res_data = empty_df
        # 清除上傳紀錄，讓使用者可以再次上傳同一個檔案
        if 'last_uploaded_id' in st.session_state:
            del st.session_state['last_uploaded_id']
        st.rerun()

st.divider()

# ================= 檔案生成與下載 =================
data_docx = create_docx(sch_year, sch_term, issue_unit, edited_df)

if data_docx:
    col_word, col_pdf = st.columns(2)
    
    with col_word:
        st.markdown("#### 🔹 第一步：下載 Word (可修改)")
        st.download_button(
            label="📥 下載 Word 檔",
            data=data_docx,
            file_name=f"正德調代課單_{datetime.date.today().strftime('%Y%m%d')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
        
    with col_pdf:
        st.markdown("#### 🔹 第二步：轉換成 PDF (手機專用)")
        if st.button("🔄 呼叫雲端伺服器轉成 PDF", use_container_width=True):
            with st.spinner("🚀 伺服器正在努力轉換中 (約需 5~10 秒，請耐心等候)..."):
                pdf_data = docx_to_pdf(data_docx)
                if pdf_data:
                    st.session_state['ready_pdf'] = pdf_data
                    st.success("✅ 轉換成功！請點擊下方按鈕下載。")
                else:
                    st.error("❌ 轉換失敗，伺服器過度繁忙或缺少套件。")
                    
        # 若有轉換好的 PDF，則顯示下載按鈕
        if 'ready_pdf' in st.session_state:
            st.download_button(
                label="📥 點我下載生成的 PDF 檔",
                data=st.session_state['ready_pdf'],
                file_name=f"正德調代課單_{datetime.date.today().strftime('%Y%m%d')}.pdf",
                mime="application/pdf",
                type="primary",
                use_container_width=True
            )