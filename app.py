import streamlit as st
import pandas as pd
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, Alignment
from PIL import Image as PILImage, ImageOps
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from openpyxl.cell import MergedCell
import os
from io import BytesIO
import math

# --- Streamlit UI 設定 ---
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
    # recorder 變數儲存記錄人姓名
    recorder = st.text_input("記錄人")

# --- 人員配置 ---
st.header("👥 人力配置")
st.write("請填寫供應商人員與外包人員的分類人數")
role_types = ["機械", "電機", "土木", "軟體"]
staff_data = {}

# 供應商人員輸入
cols_sup = st.columns(len(role_types) + 1)
cols_sup[0].markdown("#### 供應商人員")
staff_data['供應商人員'] = []
for i, role in enumerate(role_types):
    count = cols_sup[i+1].number_input(f"供應商-{role}", min_value=0, step=1, key=f"sup_{role}")
    staff_data['供應商人員'].append(count)

# 外包人員輸入
cols_sub = st.columns(len(role_types) + 1)
cols_sub[0].markdown("#### 外包人員")
staff_data['外包人員'] = []
for i, role in enumerate(role_types):
    count = cols_sub[i+1].number_input(f"外包-{role}", min_value=0, step=1, key=f"sub_{role}")
    staff_data['外包人員'].append(count)


# --- 裝機進度 ---
# 初始化 session state
if "machine_sections" not in st.session_state:
    st.session_state["machine_sections"] = []

st.header("🏗️ 裝機進度紀錄")
new_machine_name = st.text_input("輸入新機台名稱", key="new_machine_input")
add_machine_button = st.button("➕ 新增機台")

# 添加新機台到 session state
if add_machine_button and new_machine_name:
    if new_machine_name not in st.session_state["machine_sections"]:
        st.session_state["machine_sections"].append(new_machine_name)
        st.success(f"已新增機台: {new_machine_name}")
        # 清空輸入框可能需要 st.rerun() 或更複雜的狀態管理

# 顯示每個機台的輸入欄位
progress_entries = []
# 使用 enumerate 為 key 提供唯一性
for idx, machine_name in enumerate(st.session_state["machine_sections"]):
    with st.expander(f"🔧 {machine_name} (點此展開/收合)", expanded=True):
        # 每個機台最多 4 項紀錄
        for i in range(1, 5):
            st.markdown(f"**第 {i} 項**")
            cols = st.columns([4, 1, 2]) # 內容, 人力, 備註
            # 組合 key 確保唯一
            content = cols[0].text_input(f"內容", key=f"machine_{idx}_content_{i}")
            manpower = cols[1].number_input(f"人力", key=f"machine_{idx}_manpower_{i}", min_value=0, step=1)
            note = cols[2].text_input(f"備註", key=f"machine_{idx}_note_{i}")
            # 只有當內容不為空時才添加到列表
            if content:
                progress_entries.append([machine_name, i, content, manpower, note])

# --- 週邊工作 ---
st.header("🔧 週邊工作紀錄（最多 6 項）")
side_entries = []
# 項目 1 到 6
for i in range(1, 7):
    st.markdown(f"**第 {i} 項**")
    cols = st.columns([4, 1, 2]) # 內容, 人力, 備註
    # 使用唯一 key
    content = cols[0].text_input(f"內容 ", key=f"side_content_{i}") # 空格避免與上面重複? 最好確認一下
    manpower = cols[1].number_input(f"人力 ", key=f"side_manpower_{i}", min_value=0, step=1)
    note = cols[2].text_input(f"備註 ", key=f"side_note_{i}")
    # 只記錄有內容的項目
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

# --- 點擊按鈕，開始產生 Excel ---
if st.button("✅ 產出 Excel"):

    # 創建 Excel 工作簿和工作表
    wb = Workbook()
    ws = wb.active
    ws.title = "安裝日記"

    # --- 定義 Excel 樣式 ---
    bold_font = Font(name="標楷體", size=11, bold=True)
    normal_font = Font(name="標楷體", size=11)
    thin_border_side = Side(style='thin', color='000000')
    thin_border = Border(
        left=thin_border_side,
        right=thin_border_side,
        top=thin_border_side,
        bottom=thin_border_side
    )
    # 置中對齊 + 自動換行
    center_align_wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)
    # 靠左對齊 + 自動換行
    left_align_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # --- 設定固定欄寬和預設列高 ---
    DEFAULT_COL_WIDTH = 18  # A-F 欄寬度
    DEFAULT_ROW_HEIGHT = 25 # 一般文字列高度
    IMAGE_ROW_HEIGHT = 120  # 圖片列高度 (可調整)
    NUM_COLS_TOTAL = 6      # 使用 A 到 F 欄

    # 設定欄寬
    for i in range(1, NUM_COLS_TOTAL + 1):
        ws.column_dimensions[get_column_letter(i)].width = DEFAULT_COL_WIDTH

    # 初始化目前寫入的列數
    current_row = 1

    # --- 輔助函數：寫入儲存格值並套用樣式 ---
    def write_styled_cell(row, col, value, font, alignment, border=thin_border):
        """寫入值和樣式到指定儲存格，主要用於非合併或合併區左上角。"""
        cell = ws.cell(row=row, column=col)
        # 假設呼叫此函數時會正確處理合併儲存格的值寫入(只寫左上角)
        cell.value = value
        cell.font = font
        cell.alignment = alignment
        if border:
            cell.border = border
        # 確保列有預設高度
        current_height = ws.row_dimensions[row].height
        if current_height is None or current_height < DEFAULT_ROW_HEIGHT:
             ws.row_dimensions[row].height = DEFAULT_ROW_HEIGHT

    # --- 輔助函數：僅應用樣式到儲存格 ---
    def apply_styles_only(row, col, font, alignment, border=thin_border):
        """僅應用樣式到指定儲存格，用於合併區的其他儲存格。"""
        cell = ws.cell(row=row, column=col)
        # 不設定 cell.value
        cell.font = font
        cell.alignment = alignment
        if border:
            cell.border = border
        # 確保列有預設高度
        current_height = ws.row_dimensions[row].height
        if current_height is None or current_height < DEFAULT_ROW_HEIGHT:
              ws.row_dimensions[row].height = DEFAULT_ROW_HEIGHT

    # --- 區塊 1：寫入基本資訊 ---
    # 日期
    ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=NUM_COLS_TOTAL) # B:F
    write_styled_cell(current_row, 1, "日期", bold_font, center_align_wrap) # A
    write_styled_cell(current_row, 2, str(install_date), normal_font, center_align_wrap) # B (左上角)
    for c in range(3, NUM_COLS_TOTAL + 1): apply_styles_only(current_row, c, normal_font, center_align_wrap, thin_border) # C-F 樣式
    current_row += 1
    # 天氣
    ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=NUM_COLS_TOTAL) # B:F
    write_styled_cell(current_row, 1, "天氣", bold_font, center_align_wrap) # A
    write_styled_cell(current_row, 2, weather, normal_font, center_align_wrap) # B (左上角)
    for c in range(3, NUM_COLS_TOTAL + 1): apply_styles_only(current_row, c, normal_font, center_align_wrap, thin_border) # C-F 樣式
    current_row += 1
    # 記錄人 (表頭)
    ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=NUM_COLS_TOTAL) # B:F
    write_styled_cell(current_row, 1, "記錄人", bold_font, center_align_wrap) # A
    write_styled_cell(current_row, 2, recorder, normal_font, center_align_wrap) # B (左上角) - 注意這裡顯示的是表頭的記錄人
    for c in range(3, NUM_COLS_TOTAL + 1): apply_styles_only(current_row, c, normal_font, center_align_wrap, thin_border) # C-F 樣式
    current_row += 1
    # 空一行
    ws.row_dimensions[current_row].height = DEFAULT_ROW_HEIGHT
    current_row += 1

    # --- 區塊 2：寫入人力配置 ---
    # 標題列
    header_staff = ["人員分類", *role_types, "總計"]
    for col_idx, header_text in enumerate(header_staff, 1):
        if col_idx <= NUM_COLS_TOTAL: write_styled_cell(current_row, col_idx, header_text, bold_font, center_align_wrap)
    current_row += 1
    # 數據列
    for group in ["供應商人員", "外包人員"]: # 使用正確的鍵名
        group_counts = staff_data.get(group, [])
        # 基本的數據驗證和處理
        processed_counts = []
        valid_data = True
        if isinstance(group_counts, list):
            for item in group_counts:
                if isinstance(item, (int, float)): processed_counts.append(item)
                else:
                    try: processed_counts.append(int(item))
                    except (ValueError, TypeError): valid_data = False; st.warning(f"'{group}' 數據警告..."); processed_counts.append(0)
        else: valid_data = False; st.warning(f"'{group}' 格式警告..."); processed_counts = [0] * len(role_types)
        total = sum(processed_counts) if valid_data or processed_counts else 0
        # 組合行數據
        row_data = [group, *processed_counts, total]
        # 寫入儲存格
        for col_idx, cell_value in enumerate(row_data, 1):
             if col_idx <= NUM_COLS_TOTAL:
                align = left_align_wrap if col_idx == 1 else center_align_wrap
                write_styled_cell(current_row, col_idx, cell_value, normal_font, align)
        current_row += 1 # 換行
    # 空一行
    ws.row_dimensions[current_row].height = DEFAULT_ROW_HEIGHT
    current_row += 1

    # --- 區塊 3：寫入裝機進度 ---
    if progress_entries:
        # 區塊標題
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=NUM_COLS_TOTAL) # A:F
        write_styled_cell(current_row, 1, "裝機進度", bold_font, center_align_wrap) # A (左上角)
        for c in range(2, NUM_COLS_TOTAL + 1): apply_styles_only(current_row, c, bold_font, center_align_wrap, thin_border) # B-F 樣式
        current_row += 1
        # 細項標題
        header_progress = ["機台", "項次", "內容", "人力", "備註"]
        write_styled_cell(current_row, 1, header_progress[0], bold_font, center_align_wrap) # A
        write_styled_cell(current_row, 2, header_progress[1], bold_font, center_align_wrap) # B
        ws.merge_cells(start_row=current_row, start_column=3, end_row=current_row, end_column=4) # C:D
        write_styled_cell(current_row, 3, header_progress[2], bold_font, center_align_wrap) # C (左上角)
        apply_styles_only(current_row, 4, bold_font, center_align_wrap, thin_border) # D 樣式
        write_styled_cell(current_row, 5, header_progress[3], bold_font, center_align_wrap) # E
        write_styled_cell(current_row, 6, header_progress[4], bold_font, center_align_wrap) # F
        current_row += 1
        # 數據列
        for row_data in progress_entries:
            machine, item, content, manpower, note = row_data
            write_styled_cell(current_row, 1, machine, normal_font, left_align_wrap) # A
            write_styled_cell(current_row, 2, item, normal_font, center_align_wrap)    # B
            ws.merge_cells(start_row=current_row, start_column=3, end_row=current_row, end_column=4) # C:D
            write_styled_cell(current_row, 3, content, normal_font, left_align_wrap)  # C (左上角)
            apply_styles_only(current_row, 4, normal_font, left_align_wrap, thin_border) # D 樣式
            write_styled_cell(current_row, 5, manpower, normal_font, center_align_wrap) # E
            write_styled_cell(current_row, 6, note, normal_font, left_align_wrap)     # F
            current_row += 1
        # 空一行
        ws.row_dimensions[current_row].height = DEFAULT_ROW_HEIGHT
        current_row += 1

    # --- 區塊 4：寫入週邊工作 ---
    if side_entries:
        # 區塊標題
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=NUM_COLS_TOTAL) # A:F
        write_styled_cell(current_row, 1, "週邊工作", bold_font, center_align_wrap) # A (左上角)
        for c in range(2, NUM_COLS_TOTAL + 1): apply_styles_only(current_row, c, bold_font, center_align_wrap, thin_border) # B-F 樣式
        current_row += 1
        # 細項標題
        header_side = ["項次", "內容", "人力", "備註"]
        write_styled_cell(current_row, 1, header_side[0], bold_font, center_align_wrap) # A
        ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=4) # B:D
        write_styled_cell(current_row, 2, header_side[1], bold_font, center_align_wrap) # B (左上角)
        apply_styles_only(current_row, 3, bold_font, center_align_wrap, thin_border) # C 樣式
        apply_styles_only(current_row, 4, bold_font, center_align_wrap, thin_border) # D 樣式
        write_styled_cell(current_row, 5, header_side[2], bold_font, center_align_wrap) # E
        write_styled_cell(current_row, 6, header_side[3], bold_font, center_align_wrap) # F
        current_row += 1
        # 數據列
        for row_data in side_entries:
            item, content, manpower, note = row_data
            write_styled_cell(current_row, 1, item, normal_font, center_align_wrap) # A
            ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=4) # B:D
            write_styled_cell(current_row, 2, content, normal_font, left_align_wrap) # B (左上角)
            apply_styles_only(current_row, 3, normal_font, left_align_wrap, thin_border) # C 樣式
            apply_styles_only(current_row, 4, normal_font, left_align_wrap, thin_border) # D 樣式
            write_styled_cell(current_row, 5, manpower, normal_font, center_align_wrap) # E
            write_styled_cell(current_row, 6, note, normal_font, left_align_wrap)    # F
            current_row += 1
        # 空一行
        ws.row_dimensions[current_row].height = DEFAULT_ROW_HEIGHT
        current_row += 1

    # --- 區塊 5：處理圖片區域 ---
    if photos:
        # 圖片區前的空行
        ws.row_dimensions[current_row].height = DEFAULT_ROW_HEIGHT
        current_row += 1
        # 圖片區標題
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=NUM_COLS_TOTAL) # A:F
        write_styled_cell(current_row, 1, "進度留影", bold_font, center_align_wrap, border=None) # A (左上角), 無邊框
        for c in range(2, NUM_COLS_TOTAL + 1): apply_styles_only(current_row, c, bold_font, center_align_wrap, border=None) # B-F 樣式, 無邊框
        ws.row_dimensions[current_row].height = DEFAULT_ROW_HEIGHT # 設置標題列高度
        current_row += 1

        # 計算圖片目標尺寸 (保持不變)
        try: default_char_width_approx = 7; target_img_width_px = int(DEFAULT_COL_WIDTH * 3 * default_char_width_approx)
        except: target_img_width_px = int(18 * 3 * 7)
        target_img_height_pt = IMAGE_ROW_HEIGHT - 10; target_img_height_px = int(target_img_height_pt / 0.75)

        # 圖片排列設定
        img_col_width = 3 # 每圖佔 3 欄
        num_img_cols = 2  # 每列放 2 圖

        # 遍歷照片並放置
        for i in range(0, len(photos), num_img_cols):
            # 設定圖片列和說明列的高度
            ws.row_dimensions[current_row].height = IMAGE_ROW_HEIGHT      # 圖片列
            ws.row_dimensions[current_row + 1].height = DEFAULT_ROW_HEIGHT  # 說明列

            # 處理該行的圖片 (最多 num_img_cols 張)
            for j in range(num_img_cols):
                photo_index = i + j
                if photo_index < len(photos): # 確保索引有效
                    img_file = photos[photo_index]
                    filename = img_file.name

                    try:
                        # 讀取、校正方向、縮放圖片
                        img = PILImage.open(img_file)
                        img = ImageOps.exif_transpose(img)
                        img_w, img_h = img.size
                        if img_w == 0 or img_h == 0: raise ValueError("圖片寬高為0")
                        ratio = min(target_img_width_px / img_w, target_img_height_px / img_h)
                        if ratio < 1.0:
                            new_w = int(img_w * ratio); new_h = int(img_h * ratio)
                            if new_w > 0 and new_h > 0: img_resized = img.resize((new_w, new_h), PILImage.Resampling.LANCZOS)
                            else: img_resized = img # 縮放尺寸無效則使用原圖
                        else: img_resized = img # 不放大

                        # 將圖片存入記憶體緩衝區
                        img_buffer = BytesIO()
                        img_resized.save(img_buffer, format='PNG')
                        img_buffer.seek(0)

                        # 計算圖片放置位置
                        col_start = 1 + j * img_col_width # A=1, D=4, ...
                        anchor_cell = f"{get_column_letter(col_start)}{current_row}"

                        # 添加圖片到 Excel
                        xl_img = XLImage(img_buffer)
                        ws.add_image(xl_img, anchor_cell)

                        # 合併圖片下方的說明儲存格並寫入文字
                        col_end = col_start + img_col_width - 1 # C=3, F=6, ...
                        merge_range_caption = f"{get_column_letter(col_start)}{current_row + 1}:{get_column_letter(col_end)}{current_row + 1}"
                        ws.merge_cells(merge_range_caption)
                        write_styled_cell(current_row + 1, col_start, f"說明：{filename}", normal_font, center_align_wrap) # 左上角
                        for c_idx in range(col_start + 1, col_end + 1): apply_styles_only(current_row + 1, c_idx, normal_font, center_align_wrap, thin_border) # 其他部分樣式

                        # 為圖片所在的儲存格區域添加邊框 (覆蓋在圖片下方)
                        for r_idx in [current_row]:
                            for c_idx in range(col_start, col_end + 1):
                                # 不設定值，僅應用邊框和垂直對齊
                                apply_styles_only(r_idx, c_idx, normal_font, Alignment(vertical="center"), thin_border)

                    except Exception as e:
                        # 錯誤處理：在說明區顯示錯誤訊息
                        st.error(f"處理圖片 {filename} 時發生錯誤: {e}")
                        col_start = 1 + j * img_col_width; col_end = col_start + img_col_width - 1
                        merge_range_caption = f"{get_column_letter(col_start)}{current_row + 1}:{get_column_letter(col_end)}{current_row + 1}"
                        try: ws.merge_cells(merge_range_caption) # 嘗試合併
                        except: pass # 合併失敗就算了
                        write_styled_cell(current_row + 1, col_start, f"圖片錯誤", normal_font, center_align_wrap) # 寫入左上角
                        for c_idx in range(col_start + 1, col_end + 1): apply_styles_only(current_row + 1, c_idx, normal_font, center_align_wrap, thin_border) # 其他部分樣式

            # 處理完一行圖片（含說明），行數加 2
            current_row += 2

    # --- 區塊 6：添加記錄人資訊 (始終在最下方) ---
    # 添加一個空行分隔
    ws.row_dimensions[current_row].height = DEFAULT_ROW_HEIGHT
    current_row += 1

    # 準備記錄人文字
    recorder_text = f"記錄人： {recorder}" # 使用從 UI 獲取的 recorder 變數
    merge_start_col = 1
    merge_end_col = 3 # 合併 A 到 C
    merge_range_recorder = f"{get_column_letter(merge_start_col)}{current_row}:{get_column_letter(merge_end_col)}{current_row}"

    # 嘗試合併儲存格
    try:
        ws.merge_cells(merge_range_recorder)
    except Exception as merge_err:
         st.warning(f"合併記錄人儲存格時出錯: {merge_err}. 將只寫入 A 欄。")
         merge_end_col = merge_start_col # Fallback

    # 寫入記錄人文字到左上角儲存格 (無邊框)
    write_styled_cell(current_row, merge_start_col, recorder_text, normal_font, left_align_wrap, border=None)

    # 為合併區域的其他部分應用樣式 (無邊框)
    if merge_end_col > merge_start_col:
        for c in range(merge_start_col + 1, merge_end_col + 1):
             apply_styles_only(current_row, c, normal_font, left_align_wrap, border=None)

    # --- 區塊 7：儲存與下載 ---
    # 將 Excel 存入記憶體緩衝區
    excel_file = BytesIO()
    wb.save(excel_file)
    excel_file.seek(0) # 將指針移回開頭

    # 準備下載
    file_name = f"安裝日記_{install_date}.xlsx"
    st.download_button(
        label="📥 下載 Excel 檔案",
        data=excel_file, # 從記憶體提供數據
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # 顯示成功訊息
    st.success(f"檔案 {file_name} 已成功產生！")

# --- Script End ---
