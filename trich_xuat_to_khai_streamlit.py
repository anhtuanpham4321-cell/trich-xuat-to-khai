import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re
import os
import tempfile
import io

# ==================================================
# CAU HINH
# ==================================================
st.set_page_config(page_title="Cong cu trich xuat to khai hai quan", page_icon="📋", layout="wide")

# ==================================================
# HAM CHUYEN DOI SO
# ==================================================
def parse_vn_number(val):
    if val is None:
        return None
    if isinstance(val, float) and pd.isna(val):
        return None
    s = str(val).strip()
    if s == '' or s == '-':
        return None
    s = s.replace('.', '').replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return None

# ==================================================
# HAM DOC VA TRICH XUAT
# ==================================================
def extract_customs_data(input_path):
    try:
        df = pd.read_excel(input_path, sheet_name='TKX', header=None)
    except Exception as e:
        return None, None, [], f"Khong doc duoc file: {e}"

    so_to_khai = None
    ten_cong_ty = None

    for i in range(min(80, df.shape[0])):
        row = df.iloc[i]
        col2 = str(row.get(2, '')).strip() if pd.notna(row.get(2)) else ''
        if col2 == 'Số tờ khai' and pd.notna(row.get(4)):
            so_to_khai = str(row[4]).strip()
        col3 = str(row.get(3, '')).strip() if pd.notna(row.get(3)) else ''
        if 11 <= i <= 18 and col3 == 'Tên' and pd.notna(row.get(5)):
            ten_cong_ty = str(row[5]).strip()

    item_starts = []
    for i in range(df.shape[0]):
        val = df.iloc[i, 2] if df.shape[1] > 2 else None
        if pd.notna(val):
            s = str(val).strip()
            if re.match(r'^<\d+>$', s):
                item_starts.append(i)

    results = []
    for idx, start in enumerate(item_starts):
        item = {
            'stt': idx + 1,
            'ma_hs': None,
            'mo_ta': None,
            'so_luong_1': None,
            'so_luong_2': None,
            'don_vi': None,
            'don_gia': None,
            'tri_gia': None,
            'loai_tien': None,
        }
        for offset in range(12):
            r = start + offset
            if r >= df.shape[0]:
                break
            row = df.iloc[r]
            col2 = str(row.get(2, '')).strip() if pd.notna(row.get(2)) else ''
            col14 = str(row.get(14, '')).strip() if pd.notna(row.get(14)) else ''
            if col2 == 'Mã số hàng hóa' and pd.notna(row.get(5)):
                item['ma_hs'] = str(row[5]).strip()
            if col2 == 'Mô tả hàng hóa' and pd.notna(row.get(5)):
                item['mo_ta'] = str(row[5]).strip()
            if col14 == 'Số lượng (1)':
                if pd.notna(row.get(16)):
                    item['so_luong_1'] = parse_vn_number(row[16])
                if pd.notna(row.get(24)):
                    item['don_vi'] = str(row[24]).strip()
            if col14 == 'Số lượng (2)':
                if pd.notna(row.get(16)):
                    item['so_luong_2'] = parse_vn_number(row[16])
            if col2 == 'Trị giá hóa đơn':
                if pd.notna(row.get(5)):
                    item['tri_gia'] = parse_vn_number(row[5])
                if col14 == 'Đơn giá hóa đơn' and pd.notna(row.get(17)):
                    item['don_gia'] = parse_vn_number(row[17])
                if pd.notna(row.get(22)):
                    item['loai_tien'] = str(row[22]).strip()
        results.append(item)

    return so_to_khai, ten_cong_ty, results, None

# ==================================================
# HAM GHI FILE EXCEL
# ==================================================
def write_to_excel(so_to_khai, ten_cong_ty, results):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    headers = ["STT", "Mô tả hàng hóa", "Số lượng (1)", "Số lượng (2)", "ĐƠN VỊ TÍNH",
               "Đơn giá hóa đơn", "Trị giá hóa đơn", "Loại tiền", "SỐ TỜ KHAI", "TÊN"]

    header_font = Font(bold=True, size=11)
    header_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = thin_border

    for row_idx, r in enumerate(results, 2):
        ws.cell(row=row_idx, column=1, value=r['stt']).border = thin_border
        ws.cell(row=row_idx, column=2, value=r['mo_ta']).border = thin_border
        ws.cell(row=row_idx, column=3, value=r['so_luong_1']).border = thin_border
        ws.cell(row=row_idx, column=4, value=r['so_luong_2']).border = thin_border
        ws.cell(row=row_idx, column=5, value=r['don_vi']).border = thin_border
        ws.cell(row=row_idx, column=6, value=r['don_gia']).border = thin_border
        ws.cell(row=row_idx, column=7, value=r['tri_gia']).border = thin_border
        ws.cell(row=row_idx, column=8, value=r['loai_tien']).border = thin_border
        ws.cell(row=row_idx, column=9, value=so_to_khai).border = thin_border
        ws.cell(row=row_idx, column=10, value=ten_cong_ty).border = thin_border
        ws.cell(row=row_idx, column=3).number_format = '#,##0.00'
        ws.cell(row=row_idx, column=4).number_format = '#,##0.00'
        ws.cell(row=row_idx, column=6).number_format = '#,##0.000'
        ws.cell(row=row_idx, column=7).number_format = '#,##0.00'

    col_widths = [6, 60, 14, 14, 14, 16, 16, 10, 18, 50]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A2'

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ==================================================
# GIAO DIEN WEB
# ==================================================
st.title("📋 Cong cu trich xuat to khai hai quan")
st.markdown("---")

st.subheader("📤 Buoc 1: Tai len file to khai (.xls)")
uploaded_files = st.file_uploader(
    "Chon cac file to khai hai quan dinh dang .xls (co the chon nhieu file cung luc)",
    type=["xls"],
    accept_multiple_files=True
)

if uploaded_files:
    st.success(f"✅ Da tai len {len(uploaded_files)} file")
    st.markdown("---")

    all_results = []
    file_outputs = []

    for uploaded_file in uploaded_files:
        st.subheader(f"📄 Xu ly file: {uploaded_file.name}")

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xls") as tmp:
            tmp.write(uploaded_file.read())
            tmp_path = tmp.name

        try:
            so_to_khai, ten_cong_ty, results, error = extract_customs_data(tmp_path)

            if error:
                st.error(f"❌ Loi: {error}")
            elif not results:
                st.warning("⚠️ Khong tim thay mat hang nao trong file")
            else:
                st.info(f"📊 So to khai: {so_to_khai} | Ten cong ty: {ten_cong_ty}")
                st.success(f"✅ Trich xuat duoc {len(results)} mat hang")

                df_result = pd.DataFrame(results)
                df_display = df_result[['stt', 'ma_hs', 'mo_ta', 'so_luong_1', 'don_vi', 'don_gia', 'tri_gia', 'loai_tien']]
                df_display.columns = ['STT', 'Ma HS', 'Mo ta', 'So luong', 'Don vi', 'Don gia', 'Tri gia', 'Loai tien']
                st.dataframe(df_display, use_container_width=True)

                total_tri_gia = sum(r['tri_gia'] or 0 for r in results)
                st.metric("Tong tri gia", f"{total_tri_gia:,.2f}")

                excel_data = write_to_excel(so_to_khai, ten_cong_ty, results)
                base_name = os.path.splitext(uploaded_file.name)[0]
                output_name = f"{base_name}_KET_QUA.xlsx"

                st.download_button(
                    label="📥 Tai file ket qua Excel",
                    data=excel_data,
                    file_name=output_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                all_results.extend(results)
        except Exception as e:
            st.error(f"❌ Loi khi xu ly file: {e}")
        finally:
            os.unlink(tmp_path)

        st.markdown("---")
else:
    st.info("👆 Vui long tai len file to khai .xls de bat dau")
    st.markdown("---")
    st.subheader("💡 Huong dan:")
    st.markdown("""
    - File dau vao phai la dinh dang .xls cua he thong hai quan Viet Nam
    - File phai co sheet ten 'TKX' chua du lieu
    - Co the chon nhieu file cung luc de xu ly song song
    - Ket qua se duoc hien thi tren web va co the tai ve dang Excel
    """)
#（注：内容由AI生成）
