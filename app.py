import streamlit as st
import pandas as pd
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, Alignment
from PIL import Image as PILImage, ImageOps
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from openpyxl.cell import MergedCell # 雖然不直接用它判斷，但了解它有幫助
import os
from io import BytesIO # <--- 導入 BytesIO
import math

st.set_page_config(page_title="工廠安裝日記", layout="wide")
st.title("🛠️ 工廠安裝日記自動生成器")

# --- 基本資料欄位 ---
st.header("📅 基本資訊")
col1, col2, col3 = st.columns(3)
with col1:
    install_date = st.date_input("安裝日期", value=date.today())
with col2:
    weather = st.text_input("天氣")
with col3:
    recorder = st.text_input("記錄人")

# --- 人員配置 ---
st.header("👥 人力配置")
st.write("請填寫日商人員與外包人員的分類人數")
role_types = ["機械", "電機", "土木", "軟體"]
staff_data = {}

cols_jp = st.columns(len(role_types) + 1)
cols_jp[0].markdown("#### 日商人員")
staff_data['日商人員'] = []
for i, role in enumerate(role_types):
    count = cols_jp[i+1].number_input(f"商-{role}", min_value=0, step=1, key=f"jp_{role}")
    staff_data['日商人員'].append(count)

cols_sub = st.columns(len(role_types) + 1)
cols_sub[0].markdown("#### 外包人員")
staff_data['外包人員'] = []
for i, role in enumerate(role_types):
    count = cols_sub[i+1].number_input(f"包-{role}", min_value=0, step=1, key=f"sub_{role}")
    staff_data['外包人員'].append(count)


# --- 裝機進度 ---
if "machine_sections" not in st.session_state:
    st.session_state["machine_sections"] = []

st.header("🏗️ 裝機進度紀錄")
new_machine_name = st.text_input("輸入新機台名稱", key="new_machine_input")
add_machine_button = st.button("➕ 新增機台")

if add_machine_button and new_machine_name:
    if new_machine_name not in st.session_state["machine_sections"]:
        st.session_state["machine_sections"].append(new_machine_name)
        st.success(f"已新增機台: {new_machine_name}")

progress_entries = []
for idx, machine_name in enumerate(st.session_state["machine_sections"]):
    with st.expander(f"🔧 {machine_name} (點此展開/收合)", expanded=True):
        for i in range(1, 5):
            st.markdown(f"**第 {i} 項**")
            cols = st.columns([4, 1, 2])
            content = cols[0].text_input(f"內容", key=f"machine_{idx}_content_{i}")
            manpower = cols[1].number_input(f"人力", key=f"machine_{idx}_manpower_{i}", min_value=0, step=1)
            note = cols[2].text_input(f"備註", key=f"machine_{idx}_note_{i}")
            if content:
                progress_entries.append([machine_name, i, content, manpower, note])

# --- 週邊工作 ---
st.header("🔧 週邊工作紀錄（最多 6 項）")
side_entries = []
for i in range(1, 7):
    st.markdown(f"**第 {i} 項**")
    cols = st.columns([4, 1, 2])
    content = cols[0].text_input(f"內容", key=f"side_content_{i}")
    manpower = cols[1].number_input(f"人力", key=f"side_manpower_{i}", min_value=0, step=1)
    note = cols[2].text_input(f"備註", key=f"side_note_{i}")
    if content:
        side_entries.append([i, content, manpower, note])

# --- 照片上傳 ---
st.header("📸 上傳照片")
st.markdown("**進度留影**")
photos = st.file_uploader(
    "可多選照片（jpg/png/jpeg）",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
    key="photo_uploader"
)

# --- 產生 Excel 按鈕 ---
if st.button("✅ 產出 Excel"):
    # if not photos:
    #     st.warning("尚未上傳任何照片。確定要產生沒有照片的報告嗎？")

    wb = Workbook()
    ws = wb.active
    ws.title = "安裝日記"

    # --- 定義樣式 ---
    bold_font = Font(name="標楷體", size=11, bold=True)
    normal_font = Font(name="標楷體", size=11)
    thin_border_side = Side(style='thin', color='000000')
    thin_border = Border(
        left=thin_border_side,
        right=thin_border_side,
        top=thin_border_side,
        bottom=thin_border_side
    )
    center_align_wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # --- 設定固定欄寬和預設列高 ---
    DEFAULT_COL_WIDTH = 18
    DEFAULT_ROW_HEIGHT = 25
    IMAGE_ROW_HEIGHT = 120
    NUM_COLS_TOTAL = 6

    for i in range(1, NUM_COLS_TOTAL + 1):
        ws.column_dimensions[get_column_letter(i)].width = DEFAULT_COL_WIDTH

    current_row = 1

    # --- 輔助函數：寫入儲存格並套用樣式 (保持不變) ---
    # 這個函數現在只應該被呼叫來寫入 *非合併* 儲存格，或者合併儲存格的 *左上角* 儲存格
    def write_styled_cell(row, col, value, font, alignment, border=thin_border):
        # 獲取儲存格，如果它是 MergedCell，也沒關係，因為下面只設置樣式
        cell = ws.cell(row=row, column=col)
        # *** 重要：只在它不是 MergedCell 的一部分時才設置值（或者它是合併區的左上角）***
        # 爲了簡化，我們假設呼叫此函數時，如果目標是合併區，則一定是左上角
        # 最安全的做法是在呼叫前判斷，或者修改此函數增加 isinstance(cell, MergedCell) 判斷
        # 這裡我們先假設呼叫者會正確使用 (即只對左上角儲存格設定 value)
        cell.value = value
        cell.font = font
        cell.alignment = alignment
        if border:
            cell.border = border
        current_height = ws.row_dimensions[row].height
        if current_height is None or current_height < DEFAULT_ROW_HEIGHT:
             ws.row_dimensions[row].height = DEFAULT_ROW_HEIGHT

    # --- 輔助函數：僅應用樣式到儲存格 ---
    # 這個函數用來對合併區域內的其他儲存格應用樣式
    def apply_styles_only(row, col, font, alignment, border=thin_border):
         cell = ws.cell(row=row, column=col)
         # 不設定 cell.value
         cell.font = font
         cell.alignment = alignment
         if border:
             cell.border = border
         # 確保列高被設定
         current_height = ws.row_dimensions[row].height
         if current_height is None or current_height < DEFAULT_ROW_HEIGHT:
              ws.row_dimensions[row].height = DEFAULT_ROW_HEIGHT


    # --- 寫入基本資訊 ---
    ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=NUM_COLS_TOTAL)
    write_styled_cell(current_row, 1, "日期", bold_font, center_align_wrap)
    write_styled_cell(current_row, 2, str(install_date), normal_font, center_align_wrap) # 寫入左上角 B1
    # ******** 修改處 ********
    # 對合併區域內的其他儲存格 (C1 到 F1) 僅應用樣式
    for c in range(3, NUM_COLS_TOTAL + 1):
        apply_styles_only(current_row, c, normal_font, center_align_wrap, thin_border)
    # ***********************
    current_row += 1

    ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=NUM_COLS_TOTAL)
    write_styled_cell(current_row, 1, "天氣", bold_font, center_align_wrap)
    write_styled_cell(current_row, 2, weather, normal_font, center_align_wrap) # 寫入左上角 B2
    # ******** 修改處 ********
    for c in range(3, NUM_COLS_TOTAL + 1):
        apply_styles_only(current_row, c, normal_font, center_align_wrap, thin_border)
    # ***********************
    current_row += 1

    ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=NUM_COLS_TOTAL)
    write_styled_cell(current_row, 1, "記錄人", bold_font, center_align_wrap)
    write_styled_cell(current_row, 2, recorder, normal_font, center_align_wrap) # 寫入左上角 B3
    # ******** 修改處 ********
    for c in range(3, NUM_COLS_TOTAL + 1):
        apply_styles_only(current_row, c, normal_font, center_align_wrap, thin_border)
    # ***********************
    current_row += 1

    ws.row_dimensions[current_row].height = DEFAULT_ROW_HEIGHT # 空行
    current_row += 1

    # --- 寫入人力配置 ---
    header_staff = ["人員分類", *role_types, "總計"]
    # 人力配置標題不合併，所以直接用 write_styled_cell
    for col_idx, header_text in enumerate(header_staff, 1):
        if col_idx <= NUM_COLS_TOTAL:
            write_styled_cell(current_row, col_idx, header_text, bold_font, center_align_wrap)
    current_row += 1

    # 人力配置數據不合併
    for group in ["日商人員", "外包人員"]:
        total = sum(staff_data[group])
        row_data = [group, *staff_data[group], total]
        for col_idx, cell_value in enumerate(row_data, 1):
             if col_idx <= NUM_COLS_TOTAL:
                align = left_align_wrap if col_idx == 1 else center_align_wrap
                write_styled_cell(current_row, col_idx, cell_value, normal_font, align)
    current_row += 1

    ws.row_dimensions[current_row].height = DEFAULT_ROW_HEIGHT # 空行
    current_row += 1

    # --- 寫入裝機進度 ---
    if progress_entries:
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=NUM_COLS_TOTAL) # 合併 A 到 F
        write_styled_cell(current_row, 1, "裝機進度", bold_font, center_align_wrap) # 寫入左上角 A
        # ******** 修改處 ********
        for c in range(2, NUM_COLS_TOTAL + 1): # 對 B 到 F 僅應用樣式
            apply_styles_only(current_row, c, bold_font, center_align_wrap, thin_border)
        # ***********************
        current_row += 1

        # 細項標題列處理
        header_progress = ["機台", "項次", "內容", "人力", "備註"]
        write_styled_cell(current_row, 1, header_progress[0], bold_font, center_align_wrap) # 機台 (A)
        write_styled_cell(current_row, 2, header_progress[1], bold_font, center_align_wrap) # 項次 (B)
        ws.merge_cells(start_row=current_row, start_column=3, end_row=current_row, end_column=4) # 內容 (C+D)
        write_styled_cell(current_row, 3, header_progress[2], bold_font, center_align_wrap) # 寫入左上角 C
        # ******** 修改處 ********
        apply_styles_only(current_row, 4, bold_font, center_align_wrap, thin_border) # 對 D 僅應用樣式
        # ***********************
        write_styled_cell(current_row, 5, header_progress[3], bold_font, center_align_wrap) # 人力 (E)
        write_styled_cell(current_row, 6, header_progress[4], bold_font, center_align_wrap) # 備註 (F)
        current_row += 1

        # 裝機進度數據列處理
        for row_data in progress_entries:
            machine, item, content, manpower, note = row_data
            write_styled_cell(current_row, 1, machine, normal_font, left_align_wrap)
            write_styled_cell(current_row, 2, item, normal_font, center_align_wrap)
            ws.merge_cells(start_row=current_row, start_column=3, end_row=current_row, end_column=4) # 合併內容 (C+D)
            write_styled_cell(current_row, 3, content, normal_font, left_align_wrap) # 寫入左上角 C
            # ******** 修改處 ********
            apply_styles_only(current_row, 4, normal_font, left_align_wrap, thin_border) # 對 D 僅應用樣式
            # ***********************
            write_styled_cell(current_row, 5, manpower, normal_font, center_align_wrap)
            write_styled_cell(current_row, 6, note, normal_font, left_align_wrap)
            current_row += 1

        ws.row_dimensions[current_row].height = DEFAULT_ROW_HEIGHT # 空行
        current_row += 1

    # --- 寫入週邊工作 ---
    if side_entries:
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=NUM_COLS_TOTAL) # 合併 A 到 F
        write_styled_cell(current_row, 1, "週邊工作", bold_font, center_align_wrap) # 寫入左上角 A
        # ******** 修改處 ********
        for c in range(2, NUM_COLS_TOTAL + 1): # 對 B 到 F 僅應用樣式
            apply_styles_only(current_row, c, bold_font, center_align_wrap, thin_border)
        # ***********************
        current_row += 1

        # 細項標題列處理
        header_side = ["項次", "內容", "人力", "備註"]
        write_styled_cell(current_row, 1, header_side[0], bold_font, center_align_wrap) # 項次 (A)
        ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=4) # 內容 (B+C+D)
        write_styled_cell(current_row, 2, header_side[1], bold_font, center_align_wrap) # 寫入左上角 B
        # ******** 修改處 ********
        apply_styles_only(current_row, 3, bold_font, center_align_wrap, thin_border) # 對 C 僅應用樣式
        apply_styles_only(current_row, 4, bold_font, center_align_wrap, thin_border) # 對 D 僅應用樣式
        # ***********************
        write_styled_cell(current_row, 5, header_side[2], bold_font, center_align_wrap) # 人力 (E)
        write_styled_cell(current_row, 6, header_side[3], bold_font, center_align_wrap) # 備註 (F)
        current_row += 1

        # 週邊工作數據列處理
        for row_data in side_entries:
            item, content, manpower, note = row_data
            write_styled_cell(current_row, 1, item, normal_font, center_align_wrap)
            ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=4) # 合併內容 (B+C+D)
            write_styled_cell(current_row, 2, content, normal_font, left_align_wrap) # 寫入左上角 B
            # ******** 修改處 ********
            apply_styles_only(current_row, 3, normal_font, left_align_wrap, thin_border) # 對 C 僅應用樣式
            apply_styles_only(current_row, 4, normal_font, left_align_wrap, thin_border) # 對 D 僅應用樣式
            # ***********************
            write_styled_cell(current_row, 5, manpower, normal_font, center_align_wrap)
            write_styled_cell(current_row, 6, note, normal_font, left_align_wrap)
            current_row += 1

        ws.row_dimensions[current_row].height = DEFAULT_ROW_HEIGHT # 空行
        current_row += 1

    # --- 處理圖片區域 ---
    if photos:
        ws.row_dimensions[current_row].height = DEFAULT_ROW_HEIGHT # 分隔空行
        current_row += 1

        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=NUM_COLS_TOTAL) # 合併圖片標題 A 到 F
        # 對標題列左上角應用樣式，不加邊框
        write_styled_cell(current_row, 1, "進度留影", bold_font, center_align_wrap, border=None)
        # ******** 修改處 ********
        # 對合併區域的其他儲存格 (B 到 F) 也應用樣式且不加邊框
        for c in range(2, NUM_COLS_TOTAL + 1):
            apply_styles_only(current_row, c, bold_font, center_align_wrap, border=None)
        # ***********************
        ws.row_dimensions[current_row].height = DEFAULT_ROW_HEIGHT
        current_row += 1

        # 計算圖片大小 (保持不變)
        try:
            default_char_width_approx = 7
            target_img_width_px = int(DEFAULT_COL_WIDTH * 3 * default_char_width_approx)
        except:
            target_img_width_px = int(18 * 3 * 7)

        target_img_height_pt = IMAGE_ROW_HEIGHT - 10
        target_img_height_px = int(target_img_height_pt / 0.75)

        # 開始放置圖片 (保持不變，因為內部的合併是在說明列，處理邏輯已包含)
        img_col_width = 3
        num_img_cols = 2

        for i in range(0, len(photos), num_img_cols):
            ws.row_dimensions[current_row].height = IMAGE_ROW_HEIGHT      # 圖片列
            ws.row_dimensions[current_row + 1].height = DEFAULT_ROW_HEIGHT  # 說明列

            for j in range(num_img_cols):
                photo_index = i + j
                if photo_index < len(photos):
                    img_file = photos[photo_index]
                    filename = img_file.name

                    try:
                        img = PILImage.open(img_file)
                        img = ImageOps.exif_transpose(img)
                        img_w, img_h = img.size
                        if img_w == 0 or img_h == 0: raise ValueError("圖片寬度或高度為 0")

                        ratio = min(target_img_width_px / img_w, target_img_height_px / img_h)
                        if ratio < 1.0:
                            new_w = int(img_w * ratio)
                            new_h = int(img_h * ratio)
                            if new_w > 0 and new_h > 0: img_resized = img.resize((new_w, new_h), PILImage.Resampling.LANCZOS)
                            else: img_resized = img
                        else: img_resized = img

                        img_buffer = BytesIO()
                        img_resized.save(img_buffer, format='PNG')
                        img_buffer.seek(0)

                        col_start = 1 + j * img_col_width
                        anchor_cell = f"{get_column_letter(col_start)}{current_row}"

                        xl_img = XLImage(img_buffer)
                        ws.add_image(xl_img, anchor_cell)

                        col_end = col_start + img_col_width - 1
                        merge_range_caption = f"{get_column_letter(col_start)}{current_row + 1}:{get_column_letter(col_end)}{current_row + 1}"
                        ws.merge_cells(merge_range_caption) # 合併說明列儲存格
                        # 寫入左上角說明文字
                        write_styled_cell(current_row + 1, col_start, f"說明：{filename}", normal_font, center_align_wrap)
                        # ******** 修改處 ********
                        # 對說明列合併區域的其他儲存格 (如果有的話) 僅應用樣式
                        for c_idx in range(col_start + 1, col_end + 1):
                             apply_styles_only(current_row + 1, c_idx, normal_font, center_align_wrap, thin_border)
                        # ***********************

                        # 為圖片所在的儲存格區域添加邊框 (保持不變)
                        for r_idx in [current_row]:
                            for c_idx in range(col_start, col_end + 1):
                                cell = ws.cell(row=r_idx, column=c_idx)
                                # 不需要設定 value，但要確保應用樣式
                                apply_styles_only(r_idx, c_idx, normal_font, Alignment(vertical="center"), thin_border)


                    except Exception as e:
                        st.error(f"處理圖片 {filename} 時發生錯誤: {e}")
                        col_start = 1 + j * img_col_width
                        col_end = col_start + img_col_width - 1
                        merge_range_caption = f"{get_column_letter(col_start)}{current_row + 1}:{get_column_letter(col_end)}{current_row + 1}"
                        ws.merge_cells(merge_range_caption)
                        write_styled_cell(current_row + 1, col_start, f"圖片錯誤: {filename}", normal_font, center_align_wrap)
                        # ******** 修改處 ********
                        for c_idx in range(col_start + 1, col_end + 1):
                             apply_styles_only(current_row + 1, c_idx, normal_font, center_align_wrap, thin_border)
                        # ***********************

            current_row += 2

    # --- 儲存與下載 ---
    excel_file = BytesIO()
    wb.save(excel_file)
    excel_file.seek(0)

    file_name = f"安裝日記_{install_date}.xlsx"

    st.download_button(
        label="📥 下載 Excel 檔案",
        data=excel_file,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.success(f"檔案 {file_name} 已成功產生！")
