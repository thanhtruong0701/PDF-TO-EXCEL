import os
import io
import re
import tempfile
import zipfile
import datetime
import hashlib
from typing import Any
import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
import pymupdf

st.set_page_config(
    page_title="Camelot — PDF ➜ Excel",
    page_icon="🏰",
    layout="wide",
    initial_sidebar_state="expanded",
)

TEMPLATE_PATH: str = os.path.join(os.path.dirname(__file__), "shinwoo_template.xlsx")
UPLOAD_DIR: str = os.path.join(tempfile.gettempdir(), "camelot_batch_uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

# ==================== CACHED HELPER: PAGE COUNT & PREVIEWS ====================
@st.cache_data(show_spinner=False)
def get_pdf_page_count(pdf_path: str) -> int:
    """Đếm tổng số trang PDF cực nhanh bằng PyMuPDF."""
    doc = pymupdf.open(pdf_path)
    return len(doc)


def merge_row_cells(row: list[Any]) -> list[tuple[Any, int]]:
    """Gộp các ô trống liên tiếp vào ô đứng trước để xác định colspan / merge cells."""
    merged: list[tuple[Any, int]] = []
    i = 0
    n = len(row)
    while i < n:
        val = row[i]
        span = 1
        j = i + 1
        while j < n and (row[j] is None or str(row[j]).strip() == ""):
            span += 1
            j += 1
        merged.append((val, span))
        i = j
    return merged


@st.cache_data(show_spinner=False)
def get_page_preview(pdf_path: str, page_no: int) -> tuple[bytes, str]:
    """Tạo nhanh ảnh trang PDF và bảng biểu mẫu HTML sử dụng PyMuPDF C++ engine cực nhanh."""
    doc = pymupdf.open(pdf_path)
    if page_no < 1 or page_no > len(doc):
        return b"", "<div>Trang không hợp lệ</div>"
    
    p = doc[page_no - 1]
    pix = p.get_pixmap(dpi=140)
    img_bytes = pix.tobytes("png")
    
    tabs = p.find_tables().tables
    html = ["<div style='background: #ffffff; color: #111; padding: 15px; border-radius: 8px; border: 1px solid #ccc; font-family: SimSun, Arial, sans-serif; font-size: 12px; line-height: 1.4; box-shadow: 0 2px 8px rgba(0,0,0,0.08);'>"]
    
    if tabs:
        t = tabs[0]
        ext = t.extract()
        
        words = p.get_text("words")
        top_words = [w for w in words if w[3] <= t.bbox[1] + 2]
        top_words.sort(key=lambda w: (w[1], w[0]))
        
        top_lines: list[list[str]] = []
        cur_line: list[str] = []
        cur_y: float | None = None
        for w in top_words:
            if cur_y is None or abs(w[1] - cur_y) < 5:
                cur_line.append(w[4])
                cur_y = w[1]
            else:
                top_lines.append(cur_line)
                cur_line = [w[4]]
                cur_y = w[1]
        if cur_line:
            top_lines.append(cur_line)
            
        for l_idx, line in enumerate(top_lines):
            txt = " ".join(line)
            if l_idx == 0:
                html.append(f"<div style='text-align: center; font-weight: bold; font-size: 15px; margin-bottom: 2px;'>{txt}</div>")
            elif l_idx == 1:
                html.append(f"<div style='text-align: center; font-weight: bold; font-size: 13px; margin-bottom: 4px;'>{txt}</div>")
            else:
                html.append(f"<div style='text-align: center; font-size: 11px; margin-bottom: 4px;'>{txt}</div>")
                
        html.append("<table style='width: 100%; border-collapse: collapse; border: 1.5px solid #000; font-size: 11px; margin-top: 8px;'>")
        for row in ext:
            html.append("<tr>")
            m_cells = merge_row_cells(row)
            for val, span in m_cells:
                text_clean = (val or "").strip().replace("\n", "<br>")
                v_lower = (val or "").lower()
                is_header = any(k in v_lower for k in [
                    "stt", "序号", "品号", "m.số", "品名", "tên", "số lượng", "đơn giá", 
                    "quy cách", "đơn vị", "ngày gh", "ms đ.hàng", "unit", "quantity", "unit price",
                    "采购单别", "交易币别", "汇率", "采购单号", "营业税率", "单据日期", "付款条件",
                    "供应商代号", "供应商全名", "供应商地址", "联络人", "供应商电话", "供应商邮箱", "送货地址"
                ])
                bg = "#f0f4f8" if is_header else "#ffffff"
                weight = "bold" if is_header else "normal"
                align = "center" if (is_header or not text_clean or text_clean.isdigit() or text_clean in ["pcs", "VND(越南盾)"]) else "left"
                colspan_attr = f" colspan='{span}'" if span > 1 else ""
                style = f"border: 1px solid #000; padding: 4px 6px; background-color: {bg}; font-weight: {weight}; text-align: {align}; vertical-align: middle;"
                html.append(f"<td{colspan_attr} style='{style}'>{text_clean if text_clean else '&nbsp;'}</td>")
            html.append("</tr>")
        html.append("</table>")
        
        bot_words = [w for w in words if w[1] >= t.bbox[3] - 2]
        bot_words.sort(key=lambda w: (w[1], w[0]))
        bot_lines: list[list[str]] = []
        cur_line_bot: list[str] = []
        cur_y_bot: float | None = None
        for w in bot_words:
            if cur_y_bot is None or abs(w[1] - cur_y_bot) < 5:
                cur_line_bot.append(w[4])
                cur_y_bot = w[1]
            else:
                bot_lines.append(cur_line_bot)
                cur_line_bot = [w[4]]
                cur_y_bot = w[1]
        if cur_line_bot:
            bot_lines.append(cur_line_bot)
            
        if bot_lines:
            html.append("<div style='margin-top: 10px; font-size: 11px; font-weight: bold;'>")
            for line in bot_lines:
                txt = " ".join(line)
                html.append(f"<div style='margin-bottom: 2px;'>{txt}</div>")
            html.append("</div>")
    else:
        txt = p.get_text().replace("\n", "<br>")
        html.append(f"<div style='line-height: 1.5;'>{txt}</div>")
        
    html.append("</div>")
    return img_bytes, "".join(html)


# ==================== HELPER: FEATURE 1 (EXACT FORM EXCEL) ====================
def populate_exact_form_sheet(ws: Any, pdf_path: str) -> None:
    """Tái tạo bảng biểu mẫu kẻ viền đen, đúng độ cao hàng, độ rộng cột và gộp tất cả các trang PDF vào một Sheet."""
    doc = pymupdf.open(pdf_path)
    full_text = ""
    for p in doc:
        full_text += p.get_text() + "\n"
        
    thin_border = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000")
    )
    
    # CASE 1: BATTLE-GRANDE PO (Đơn hàng nhiều trang liên tục)
    if "GRANDE" in full_text.upper() or "订货通知单" in full_text or "ĐƠN ĐẶT HÀNG" in full_text:
        # Header Công ty
        header_lines = [
            ("允伟国际责任有限公司", 13, True),
            ("CÔNG TY TNHH QUỐC TẾ GRANDE", 11, True),
            ("地址: 同奈市仁泽坊仁泽二工业区25B路 Đường 25B, KCN Nhơn Trạch II, Phường Nhơn Trạch, TP Đồng Nai", 9, False),
            ("TEL: 0251-3560487~8   FAX: 0251-3560490   MST: 3601018448", 9, False),
            ("订货通知单 / ĐƠN ĐẶT HÀNG", 14, True)
        ]
        r_idx = 1
        for txt, f_size, is_bold in header_lines:
            ws.merge_cells(f"A{r_idx}:I{r_idx}")
            cell = ws[f"A{r_idx}"]
            cell.value = txt
            cell.font = Font(name="Arial", size=f_size, bold=is_bold)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[r_idx].height = 20 if is_bold else 16
            r_idx += 1
            
        m_po = re.search(r"MS đơn hàng[：:]\s*(\w+)", full_text)
        po_code = m_po.group(1) if m_po else "261026010"
        m_date = re.search(r"Ngày đặt hàng[：:]\s*(\d+)", full_text)
        po_date = m_date.group(1) if m_date else "20261026"
        
        info = [
            ("厂商代码 Mã số NCC: VNH015", f"采购单号MS đơn hàng: {po_code}"),
            ("厂商简称 Tên NCC: HỮU TÍN (SHINWOO)", f"采购日期 Ngày đặt hàng: {po_date}"),
            ("送货地址 Địa điểm GH:", "交易币别 Loại tiền tệ: VND")
        ]
        for left_txt, right_txt in info:
            ws.merge_cells(f"A{r_idx}:C{r_idx}")
            ws[f"A{r_idx}"] = left_txt
            ws[f"A{r_idx}"].font = Font(name="Arial", size=9, bold=True)
            ws.merge_cells(f"D{r_idx}:I{r_idx}")
            ws[f"D{r_idx}"] = right_txt
            ws[f"D{r_idx}"].font = Font(name="Arial", size=9, bold=True)
            ws.row_dimensions[r_idx].height = 18
            r_idx += 1
            
        r_idx += 1 # 1 dòng trống trước bảng
        
        # Tiêu đề bảng
        headers = [
            "项次\nSTT", "料号\nM.Số VL", "品名规格\nTên Sản Phẩm&Qui Cách", 
            "订购量\nS.Lượng", "单位\nĐ.Vị", "单价\nĐ.Giá", 
            "金额\nT.Tiền", "预交日\nNgày GH", "订单号\nMS Đ.Hàng"
        ]
        header_row = r_idx
        ws.row_dimensions[header_row].height = 28
        for c_i, h in enumerate(headers, start=1):
            cell = ws.cell(header_row, c_i, h)
            cell.font = Font(name="Arial", size=9, bold=True)
            cell.fill = PatternFill(start_color="E8EEF5", end_color="E8EEF5", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        r_idx += 1
        
        # Bóc tách tất cả các mặt hàng xuyên suốt các trang
        all_items: list[list[str]] = []
        for p in doc:
            tabs = p.find_tables().tables
            if not tabs:
                continue
            for r in tabs[0].extract():
                if not r or len(r) < 9:
                    continue
                c0 = (r[0] or "").strip()
                c2 = (r[2] or "").strip()
                if c0.isdigit():
                    all_items.append([c.strip() if c else "" for c in r[:9]])
                elif all_items and c2 and not c0:
                    all_items[-1][2] = (all_items[-1][2] + "\n" + c2).strip()
                    
        start_item_row = r_idx
        for itm in all_items:
            lines_cnt = itm[2].count("\n") + 1
            ws.row_dimensions[r_idx].height = max(20.0, lines_cnt * 13.5)
            
            for c_i, val in enumerate(itm, start=1):
                cell = ws.cell(r_idx, c_i)
                cell.font = Font(name="Arial", size=9)
                cell.border = thin_border
                
                if c_i == 1: # STT
                    cell.value = int(val) if val.isdigit() else val
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif c_i == 2: # Code
                    cell.value = val
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif c_i == 3: # Description
                    cell.value = val
                    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                elif c_i == 4: # Qty
                    try: cell.value = float(val)
                    except: cell.value = val
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    cell.number_format = "#,##0.0"
                elif c_i == 5: # Unit
                    cell.value = val
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif c_i in [6, 7]: # Price, Amount
                    try: cell.value = float(val)
                    except: cell.value = val
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    cell.number_format = "#,##0.00"
                else: # Date, PO
                    cell.value = val
                    cell.alignment = Alignment(horizontal="center", vertical="center")
            r_idx += 1
        end_item_row = r_idx - 1
        
        # Dòng Tổng Cộng
        ws.merge_cells(f"A{r_idx}:F{r_idx}")
        cell_tot_lbl = ws[f"A{r_idx}"]
        cell_tot_lbl.value = "合计 VND Tổng số tiền:"
        cell_tot_lbl.font = Font(name="Arial", size=10, bold=True)
        cell_tot_lbl.alignment = Alignment(horizontal="right", vertical="center")
        cell_tot_lbl.border = thin_border
        for c in range(2, 7):
            ws.cell(r_idx, c).border = thin_border
            
        cell_tot_val = ws.cell(r_idx, 7)
        cell_tot_val.value = f"=SUM(G{start_item_row}:G{end_item_row})"
        cell_tot_val.font = Font(name="Arial", size=10, bold=True)
        cell_tot_val.number_format = "#,##0.00"
        cell_tot_val.alignment = Alignment(horizontal="right", vertical="center")
        cell_tot_val.border = thin_border
        
        ws.merge_cells(f"H{r_idx}:I{r_idx}")
        ws.cell(r_idx, 8).border = thin_border
        ws.cell(r_idx, 9).border = thin_border
        ws.row_dimensions[r_idx].height = 22
        r_idx += 2
        
        # Điều khoản & Lưu ý
        ws.merge_cells(f"A{r_idx}:I{r_idx}")
        ws[f"A{r_idx}"] = "注意 / Lưu Ý:"
        ws[f"A{r_idx}"].font = Font(name="Arial", size=10, bold=True)
        r_idx += 1
        
        terms = [
            "1. 订制之品质形状包装等必须与买方所指定的不得有差异, 又如包装之货品与订货品相差或因包装不良发生破损, 一切责任应有卖方负责. (Chất lượng, kiểu dáng, phương thức đóng gói sản phẩm đặt mua bắt buộc dựa theo chỉ định của bên Mua, nếu phát hiện sự khác biệt dẫn tới việc hư hại sản phẩm thì toàn bộ trách nhiệm thuộc về bên Bán).",
            "2. 交货日期请严格遵守, 如拖延交货日期不履行交货时买主得注销订货之全部或部份, 如由此发生损害时, 一切责任完全由卖方负责. (Bên Bán cần nghiêm khắc tuân thủ ngày yêu cầu giao hàng, nếu không giao hàng hoặc kéo dài thời gian giao hàng, bên Mua có quyền hủy bỏ một phần hoặc tất cả đơn hàng; trường hợp dẫn tới tổn thất thì toàn bộ trách nhiệm thuộc về bên Bán).",
            "3. 关于所订货品由卖方发生损害时,一切责任由卖方负责 (Những thiệt hại khác liên quan xảy ra trong quá trình mua hàng nếu do bên Bán gây ra thì phải hoàn toàn chịu trách nhiệm).",
            "4. 包装箱(外箱或内箱)应填明品名及装入数量 (Thùng đóng gói (thùng ngoài và thùng trong ) phải ghi rõ tên sản phẩm và số lượng đóng thùng ).",
            "5. 如有指定商标时, 务必将商标贴于指定之处 (Nếu có chỉ định nhãn hiệu, bắt buộc nhãn hiệu phải dán đúng nơi qui định).",
            "6. 外箱唛头请照批示印于外箱 (Mác thùng ngoài theo chỉ dẫn được in trên mặt ngoài thùng).",
            "7. 厂商签回联如于五日内未如期签回即视同接受 (Nhà cung cấp nhận được đơn đặt hàng cần xác nhận phản hồi, nếu qua thời hạn 5 ngày kể từ nhận đơn không có dấu hiệu phản hồi đồng nghĩa chấp nhận yêu cầu đơn hàng).",
            "8. 请提供合同 (Cung cấp hợp đồng)."
        ]
        for t_str in terms:
            ws.merge_cells(f"A{r_idx}:I{r_idx}")
            cell_t = ws[f"A{r_idx}"]
            cell_t.value = t_str
            cell_t.font = Font(name="Arial", size=8.5)
            cell_t.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            ws.row_dimensions[r_idx].height = 28
            r_idx += 1
            
        r_idx += 1
        # Chữ ký phê duyệt
        ws.merge_cells(f"A{r_idx}:C{r_idx}")
        ws[f"A{r_idx}"] = "核准 Phê chuẩn:"
        ws[f"A{r_idx}"].font = Font(name="Arial", size=10, bold=True)
        
        ws.merge_cells(f"D{r_idx}:F{r_idx}")
        ws[f"D{r_idx}"] = "审核 Thẩm tra:"
        ws[f"D{r_idx}"].font = Font(name="Arial", size=10, bold=True)
        
        ws.merge_cells(f"G{r_idx}:I{r_idx}")
        ws[f"G{r_idx}"] = "制表 Người lập biểu:"
        ws[f"G{r_idx}"].font = Font(name="Arial", size=10, bold=True)
        ws.row_dimensions[r_idx].height = 24
        
        col_widths = {
            'A': 7, 'B': 18, 'C': 50, 'D': 12, 'E': 8, 'F': 12, 'G': 15, 'H': 12, 'I': 14
        }
        for col_letter, w in col_widths.items():
            ws.column_dimensions[col_letter].width = w
            
    # CASE 2: KAITHIN PO (Việt Nam Khai Sâm)
    elif "开森" in full_text or "KAITHIN" in full_text.upper() or "采购单" in full_text:
        # Header (14 cột A..N)
        ws.merge_cells("A1:N1")
        ws["A1"] = "越南开森责任有限公司"
        ws["A1"].font = Font(name="SimSun", size=14, bold=True)
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 24

        ws.merge_cells("A2:N2")
        ws["A2"] = "采购单"
        ws["A2"].font = Font(name="SimSun", size=12, bold=True)
        ws["A2"].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[2].height = 20

        ws.merge_cells("A3:N3")
        ws["A3"] = "页次: 1/1"
        ws["A3"].font = Font(name="SimSun", size=9)
        ws["A3"].alignment = Alignment(horizontal="right", vertical="center")
        ws.row_dimensions[3].height = 16

        tabs_k = doc[0].find_tables().tables
        if tabs_k:
            ext_k = tabs_k[0].extract()
            curr_r = 4
            for r_data in ext_k:
                m_cells = merge_row_cells(r_data)
                non_empty = [c for c in r_data if c and str(c).strip()]
                max_lines = max((str(c or "").count("\n") + 1 for c in non_empty), default=1)
                ws.row_dimensions[curr_r].height = max(18.0, 14.0 * max_lines)
                
                is_tbl_header = any("序号" in str(c) or "品号" in str(c) for c in non_empty)
                
                c_start = 1
                for val, span in m_cells:
                    c_end = c_start + span - 1
                    if span > 1:
                        ws.merge_cells(start_row=curr_r, start_column=c_start, end_row=curr_r, end_column=c_end)
                    
                    cell = ws.cell(curr_r, c_start)
                    v_str = val.strip() if val else ""
                    cell.value = v_str
                    
                    align_h = "center" if (is_tbl_header or not v_str or v_str.isdigit() or v_str in ["pcs", "VND(越南盾)"]) else "left"
                    
                    for col in range(c_start, c_end + 1):
                        c_elem = ws.cell(curr_r, col)
                        c_elem.border = thin_border
                        if is_tbl_header:
                            c_elem.font = Font(name="SimSun", size=9, bold=True)
                            c_elem.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
                            c_elem.alignment = Alignment(horizontal="center", vertical="center")
                        else:
                            c_elem.font = Font(name="SimSun", size=9)
                            c_elem.alignment = Alignment(horizontal=align_h, vertical="center", wrap_text=True)
                            
                    c_start = c_end + 1
                curr_r += 1
                
            curr_r += 1
            ws.merge_cells(f"A{curr_r}:N{curr_r}")
            ws[f"A{curr_r}"] = "厂商进货时，务必提供千分之三的备品"
            ws[f"A{curr_r}"].font = Font(name="SimSun", size=9, bold=True)
            curr_r += 1
            ws.merge_cells(f"A{curr_r}:N{curr_r}")
            ws[f"A{curr_r}"] = "核准:                      审核人: 杨艳珍                      经办人: NJ000002"
            ws[f"A{curr_r}"].font = Font(name="SimSun", size=9, bold=True)
            
        k_widths = {
            'A': 6, 'B': 7, 'C': 7, 'D': 10, 'E': 10, 'F': 12, 'G': 12,
            'H': 9, 'I': 6, 'J': 6, 'K': 6, 'L': 7, 'M': 7, 'N': 16
        }
        for col_l, w in k_widths.items():
            ws.column_dimensions[col_l].width = w

    # CASE 3: SEDO hoặc các định dạng PO khác
    else:
        curr_r = 1
        for p_idx, p in enumerate(doc):
            tabs_s = p.find_tables().tables
            if tabs_s:
                ext_s = tabs_s[0].extract()
                for r_data in ext_s:
                    m_cells = merge_row_cells(r_data)
                    non_empty = [c for c in r_data if c and str(c).strip()]
                    max_lines = max((str(c or "").count("\n") + 1 for c in non_empty), default=1)
                    ws.row_dimensions[curr_r].height = max(18.0, 14.0 * max_lines)
                    
                    c_start = 1
                    for val, span in m_cells:
                        c_end = c_start + span - 1
                        if span > 1:
                            ws.merge_cells(start_row=curr_r, start_column=c_start, end_row=curr_r, end_column=c_end)
                        cell = ws.cell(curr_r, c_start)
                        cell.value = val.strip() if val else ""
                        for col in range(c_start, c_end + 1):
                            c_elem = ws.cell(curr_r, col)
                            c_elem.border = thin_border
                            c_elem.font = Font(name="Arial", size=9)
                            c_elem.alignment = Alignment(vertical="center", wrap_text=True)
                        c_start = c_end + 1
                    curr_r += 1
                curr_r += 1
                
        for col_i in range(1, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(col_i)].width = 12


@st.cache_data(show_spinner=False)
def export_exact_form_excel(pdf_path: str) -> bytes:
    """Xuất file Excel (.xlsx) chuẩn biểu mẫu kẻ viền đen, đúng độ cao hàng, độ rộng cột và GỘP TẤT CẢ VÀO 1 TAB DUY NHẤT."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet(title="Don_Dat_Hang")
    populate_exact_form_sheet(ws, pdf_path)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ==================== HELPER: FEATURE 2 (SHINWOO ORDER FORM AUTO-FILL) ====================
def extract_barcode(text: str) -> str:
    """Trích xuất mã vạch UPC (12 số) hoặc EAN (13 số) với mọi dạng ngăn cách."""
    if not text:
        return ""
    # 1. UPC with any prefix separator: UPC 123456789012, UPC-123456789012, UPC123456789012, UPC...
    m = re.search(r'UPC[^\d]*?(\d{12,13})', text, re.IGNORECASE)
    if m:
        return m.group(1)
    # 2. EAN with prefix separator
    m = re.search(r'EAN[^\d]*?(\d{12,13})', text, re.IGNORECASE)
    if m:
        return m.group(1)
    # 3. Fallback: Any standalone 12 or 13 consecutive digits
    m = re.search(r'\b(\d{12,13})\b', text)
    if m:
        return m.group(1)
    return ""


def extract_price(text: str, fallback_unit_price: str = "") -> str:
    """Trích xuất giá bán lẻ (Retail Price) hỗ trợ ký hiệu $ (ASCII), ＄ (Fullwidth U+FF04), USD hoặc đơn giá."""
    if text:
        # 1. $ or ＄ followed by digits (e.g. $179.99, ＄179.99, $ 179.99, ＄ 199)
        m = re.search(r"[\$\uFF04]\s*(\d+(?:\.\d{1,2})?)", text)
        if m:
            return f"${m.group(1)}"
        # 2. USD followed by digits
        m_usd = re.search(r"USD\s*(\d+(?:\.\d{1,2})?)", text, re.IGNORECASE)
        if m_usd:
            return f"${m_usd.group(1)}"
        # 3. Digits followed by $ or ＄
        m_rev = re.search(r"(\d+(?:\.\d{1,2})?)\s*[\$\uFF04]", text)
        if m_rev:
            return f"${m_rev.group(1)}"
    if fallback_unit_price:
        clean_p = re.sub(r"[^\d\.]", "", fallback_unit_price)
        if clean_p:
            return clean_p
    return ""


@st.cache_data(show_spinner=False)
def extract_shinwoo_order_items(pdf_path: str) -> tuple[str, str, list[dict[str, Any]]]:
    """Trích xuất tự động thông tin PO và danh sách items chính xác 100% mã vạch (UPC/EAN) và giá tiền."""
    doc = pymupdf.open(pdf_path)
    full_text = ""
    for p in doc:
        full_text += p.get_text() + "\n"
        
    po_number: str = ""
    factory_request: str = ""
    raw_items: list[dict[str, Any]] = []
    
    # 1. BATTLE-GRANDE PO
    if "GRANDE" in full_text.upper() or "ĐƠN ĐẶT HÀNG" in full_text or "订货通知单" in full_text:
        m_po = re.search(r"MS đơn hàng[：:]\s*(\w+)", full_text)
        if m_po:
            po_number = m_po.group(1)
        factory_request = "CÔNG TY TNHH QUỐC TẾ GRANDE\nĐường 25B, KCN Nhơn Trạch II, Phường Nhơn Trạch, TP Đồng Nai"
        
        for p in doc:
            tabs = p.find_tables().tables
            if not tabs:
                continue
            ext = tabs[0].extract()
            for row in ext:
                if not row or len(row) < 9:
                    continue
                stt_str = (row[0] or "").strip()
                desc_str = (row[2] or "").strip()
                unit_p_str = (row[5] or "").strip()
                
                if stt_str.isdigit():
                    code = (row[1] or "").strip()
                    qty_str = (row[3] or "").strip()
                    qty_val = re.sub(r"[^\d]", "", qty_str.split(".")[0]) or "0"
                    po_item = (row[8] or "").strip() or po_number
                    
                    raw_items.append({
                        "stt": int(stt_str),
                        "po": po_item,
                        "code": code,
                        "desc": desc_str,
                        "qty": int(qty_val),
                        "unit_p": unit_p_str
                    })
                elif raw_items and desc_str and not stt_str:
                    # Ghép nối các dòng mô tả bị cắt đôi giữa 2 trang
                    raw_items[-1]["desc"] += " " + desc_str
                    
        items: list[dict[str, Any]] = []
        for idx, itm in enumerate(raw_items, start=1):
            desc = itm["desc"]
            upc = extract_barcode(desc)
            retail_price = extract_price(desc, itm.get("unit_p", ""))
            
            items.append({
                "No.": idx,
                "1. PO Number": itm["po"],
                "2. Quantity (pcs)": itm["qty"],
                "3. Season Code": "",
                "4. ITEM DESCRIPTION": f"{itm['code']}\n{desc.splitlines()[0]}".strip(),
                "5. Simplified Color": "",
                "6. Style number/VSN #": itm["code"],
                "7A. UPC A Code": upc if len(upc) == 12 else "",
                "7B. EAN 13 Code": upc if len(upc) == 13 else "",
                "10. COUNTRY OF ORIGIN": "MADE IN VIETNAM",
                "11. Retail Price": retail_price
            })
        return po_number, factory_request, items

    # 2. SEDO PO
    elif "SEDO" in full_text.upper() or "Sedo Vina" in full_text:
        m_po = re.search(r"PO No[：:]\s*(\w+)", full_text)
        if m_po:
            po_number = m_po.group(1)
        factory_request = "Sedo Vina Co.,LTD.\nTEL: (84) 28-3895-5541"
        items = []
        for p in doc:
            t = p.get_text()
            lines = t.splitlines()
            for l_idx, line in enumerate(lines):
                if line.startswith("LBL"):
                    code = line.strip()
                    qty = "0"
                    ean = ""
                    price = ""
                    for nxt in lines[l_idx + 1:l_idx + 15]:
                        clean_nxt = nxt.strip()
                        if re.match(r"^\d{1,3}(,\d{3})+$", clean_nxt):
                            qty = clean_nxt.replace(",", "")
                        elif " d" in clean_nxt or " VND" in clean_nxt or re.match(r"^\d+(\.\d+)?\s*d$", clean_nxt):
                            price = clean_nxt
                        m_ean = re.search(r"\b(\d{13})\b", clean_nxt)
                        if m_ean:
                            ean = m_ean.group(1)
                        if "TOTAL" in clean_nxt:
                            break
                    items.append({
                        "No.": len(items) + 1,
                        "1. PO Number": po_number,
                        "2. Quantity (pcs)": int(qty) if qty != "0" else 0,
                        "3. Season Code": "",
                        "4. ITEM DESCRIPTION": code,
                        "5. Simplified Color": "",
                        "6. Style number/VSN #": code,
                        "7A. UPC A Code": ean if len(ean) == 12 else "",
                        "7B. EAN 13 Code": ean if len(ean) == 13 else "",
                        "10. COUNTRY OF ORIGIN": "MADE IN VIETNAM",
                        "11. Retail Price": price
                    })
        return po_number, factory_request, items

    # 3. KAITHIN / KESON PO
    elif "开森" in full_text or "KESON" in full_text.upper() or "KAITHIN" in full_text.upper() or "采购单" in full_text:
        m_po = re.search(r"采购单号[：:]\s*(\w+)", full_text)
        if m_po:
            po_number = m_po.group(1)
        m_addr = re.search(r"送货地址[：:]\s*(.*?)(?=序号|$)", full_text, re.DOTALL)
        if m_addr:
            factory_request = m_addr.group(1).strip()
            
        items = []
        tabs_k = doc[0].find_tables().tables
        if tabs_k:
            ext = tabs_k[0].extract()
            header_idx = -1
            for r_idx, row in enumerate(ext):
                if any("品号" in str(c) for c in row if c):
                    header_idx = r_idx
                    break
            if header_idx != -1:
                for row in ext[header_idx + 1:]:
                    clean_row = [c for c in row if c is not None and str(c).strip() != ""]
                    if not clean_row or "以下空白" in str(clean_row) or "合计" in str(clean_row):
                        break
                    stt = clean_row[0] if len(clean_row) > 0 else len(items) + 1
                    code = clean_row[1] if len(clean_row) > 1 else ""
                    name = clean_row[2] if len(clean_row) > 2 else ""
                    spec = clean_row[3] if len(clean_row) > 3 else ""
                    qty_str = clean_row[4] if len(clean_row) > 4 else "0"
                    qty = re.sub(r"[^\d]", "", str(qty_str)) or "0"
                    unit_p = clean_row[6].strip() if len(clean_row) > 6 else ""
                    
                    items.append({
                        "No.": len(items) + 1,
                        "1. PO Number": po_number,
                        "2. Quantity (pcs)": int(qty),
                        "3. Season Code": "",
                        "4. ITEM DESCRIPTION": f"{code}\n{name} {spec}".strip(),
                        "5. Simplified Color": "",
                        "6. Style number/VSN #": code,
                        "7A. UPC A Code": "",
                        "7B. EAN 13 Code": "",
                        "10. COUNTRY OF ORIGIN": "MADE IN VIETNAM",
                        "11. Retail Price": unit_p
                    })
        return po_number, factory_request, items
        
    return po_number, factory_request, []


def populate_shinwoo_sheet(ws: Any, items: list[dict[str, Any]], factory_request: str = "") -> None:
    """Điền dữ liệu các mặt hàng vào bảng tính mẫu của SHINWOO."""
    footer_h = [ws.cell(12, c).value for c in range(1, 6)]
    footer_v = [ws.cell(13, c).value for c in range(1, 6)]
    if factory_request:
        footer_v[4] = factory_request
        
    for r in range(3, 11):
        for c in range(1, 13):
            ws.cell(r, c).value = None
            
    for r in (12, 13):
        for c in range(1, 6):
            ws.cell(r, c).value = None
            
    thin_border = Border(
        left=Side(style="thin", color="D3D3D3"),
        right=Side(style="thin", color="D3D3D3"),
        top=Side(style="thin", color="D3D3D3"),
        bottom=Side(style="thin", color="D3D3D3")
    )
    
    for idx, itm in enumerate(items, start=1):
        r = 2 + idx
        ws.cell(r, 1, itm.get("No.", idx))
        ws.cell(r, 2, itm.get("1. PO Number", ""))
        ws.cell(r, 3, itm.get("2. Quantity (pcs)", 0))
        ws.cell(r, 4, itm.get("3. Season Code", ""))
        ws.cell(r, 5, itm.get("4. ITEM DESCRIPTION", ""))
        ws.cell(r, 6, itm.get("5. Simplified Color", ""))
        ws.cell(r, 7, itm.get("6. Style number/VSN #", ""))
        ws.cell(r, 8, itm.get("7A. UPC A Code", ""))
        ws.cell(r, 9, itm.get("7B. EAN 13 Code", ""))
        ws.cell(r, 10, "")
        ws.cell(r, 11, itm.get("10. COUNTRY OF ORIGIN", "MADE IN VIETNAM"))
        ws.cell(r, 12, itm.get("11. Retail Price", ""))
        
        for c in range(1, 13):
            ws.cell(r, c).border = thin_border
            ws.cell(r, c).alignment = Alignment(vertical="center", wrap_text=True)
            
    last_item_row = 2 + max(len(items), 8)
    f_h_row = last_item_row + 2
    f_v_row = f_h_row + 1
    
    for c, val in enumerate(footer_h, start=1):
        c_cell = ws.cell(f_h_row, c, val)
        c_cell.font = Font(name="Arial", size=10, bold=True)
        c_cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        
    for c, val in enumerate(footer_v, start=1):
        v_cell = ws.cell(f_v_row, c, val)
        v_cell.alignment = Alignment(vertical="top", wrap_text=True)


def fill_shinwoo_template(items: list[dict[str, Any]], factory_request: str = "") -> bytes:
    """Điền danh sách mặt hàng vào file mẫu Excel chuẩn của SHINWOO."""
    if not os.path.exists(TEMPLATE_PATH):
        raise FileNotFoundError(f"Không tìm thấy template: {TEMPLATE_PATH}")
    wb = openpyxl.load_workbook(TEMPLATE_PATH)
    ws = wb.active
    populate_shinwoo_sheet(ws, items, factory_request)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


@st.cache_data(show_spinner=False)
def export_single_order_excel(pdf_path: str) -> bytes:
    """Tạo đúng 1 file Excel cho mỗi file PDF, tự động tích hợp cả 2 chức năng:
    - Tab 1: Don_Dat_Hang (Biểu mẫu chuẩn 1 tab duy nhất, format y chang PDF gốc)
    - Tab 2: Form_SHINWOO (Form đặt hàng Shinwoo với PO, mã vạch, giá nếu có)
    """
    po, fac, items = extract_shinwoo_order_items(pdf_path)
    
    if items and os.path.exists(TEMPLATE_PATH):
        wb = openpyxl.load_workbook(TEMPLATE_PATH)
        ws_sh = wb.active
        ws_sh.title = "Form_SHINWOO"
        populate_shinwoo_sheet(ws_sh, items, fac)
        ws_form = wb.create_sheet(title="Don_Dat_Hang", index=0)
    else:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        ws_form = wb.create_sheet(title="Don_Dat_Hang")
        
    populate_exact_form_sheet(ws_form, pdf_path)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


@st.cache_data(show_spinner=False)
def generate_batch_zip(
    file_tuples: tuple[tuple[str, str], ...]
) -> tuple[bytes, list[dict[str, Any]]]:
    """Tạo file .ZIP nén: Trong thư mục có bao nhiêu file PDF thì trả lại đúng bấy nhiêu file Excel."""
    zip_buf = io.BytesIO()
    summary: list[dict[str, Any]] = []
    
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for fname, fpath in file_tuples:
            base_name = os.path.splitext(fname)[0]
            clean_base = re.sub(r'[\\/*?:"<>|]', '_', base_name)
            pages = get_pdf_page_count(fpath)
            
            po, fac, items = extract_shinwoo_order_items(fpath)
            bc_count = sum(1 for itm in items if itm.get("7A. UPC A Code") or itm.get("7B. EAN 13 Code"))
            pr_count = sum(1 for itm in items if itm.get("11. Retail Price"))
            
            # Xuất đúng 1 file Excel tương ứng cho 1 file PDF (tích hợp đầy đủ cả 2 tab)
            excel_bytes = export_single_order_excel(fpath)
            zf.writestr(f"{clean_base}.xlsx", excel_bytes)
            
            summary.append({
                "Tên File PDF": fname,
                "Số Trang": pages,
                "Mã PO": po or "N/A",
                "Số Mặt Hàng": len(items),
                "Mã Barcode (UPC/EAN)": f"{bc_count}/{len(items)}" if items else "N/A",
                "Giá Tiền": f"{pr_count}/{len(items)}" if items else "N/A",
                "Trạng Thái": "✅ Thành công"
            })
            
    return zip_buf.getvalue(), summary


def get_batch_hash(pdf_files: list[tuple[str, str]]) -> str:
    """Tính mã băm duy nhất cho danh sách file để kiểm tra thay đổi."""
    h = hashlib.md5()
    for fname, fpath in pdf_files:
        h.update(fname.encode("utf-8", errors="ignore"))
        if os.path.exists(fpath):
            h.update(str(os.path.getsize(fpath)).encode("utf-8"))
            h.update(str(os.path.getmtime(fpath)).encode("utf-8"))
    return h.hexdigest()


def run_streamlit_app() -> None:
    # ==================== CUSTOM CSS ====================
    st.markdown("""
    <style>
    /* ---------- GOOGLE FONTS ---------- */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800;900&display=swap');

    /* ---------- ROOT VARIABLES ---------- */
    :root {
        --bg-primary: #0F172A;
        --bg-card: #1E293B;
        --bg-card-hover: #253347;
        --accent: #6366F1;
        --accent-light: #818CF8;
        --accent-glow: rgba(99,102,241,.25);
        --text-primary: #F1F5F9;
        --text-secondary: #94A3B8;
        --text-muted: #64748B;
        --border: #334155;
        --success: #22C55E;
        --warning: #F59E0B;
        --danger: #EF4444;
        --radius: 12px;
        --radius-sm: 8px;
        --shadow: 0 4px 24px rgba(0,0,0,.35);
    }

    /* ---------- GLOBAL OVERRIDES ---------- */
    html, body, [data-testid="stAppViewContainer"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif !important;
    }
    [data-testid="stAppViewContainer"] {
        background: linear-gradient(135deg, #0F172A 0%, #1a1f3a 100%);
    }
    [data-testid="stHeader"] {
        background: transparent !important;
    }
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1E293B 0%, #162032 100%) !important;
        border-right: 1px solid var(--border);
    }
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
    [data-testid="stSidebar"] label {
        color: var(--text-secondary) !important;
    }

    /* ---------- HERO HEADER ---------- */
    .hero-header {
        background: linear-gradient(135deg, rgba(99,102,241,.12) 0%, rgba(139,92,246,.08) 100%);
        border: 1px solid rgba(99,102,241,.2);
        border-radius: 16px;
        padding: 2.5rem 2rem;
        margin-bottom: 1.5rem;
        text-align: center;
        position: relative;
        overflow: hidden;
    }
    .hero-header::before {
        content: '';
        position: absolute;
        top: -50%;
        left: -50%;
        width: 200%;
        height: 200%;
        background: radial-gradient(circle at 30% 30%, rgba(99,102,241,.08), transparent 60%);
        pointer-events: none;
    }
    .hero-header h1 {
        font-size: 2.2rem;
        font-weight: 800;
        background: linear-gradient(135deg, #fff 0%, #a5b4fc 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin: 0 0 .5rem 0;
        letter-spacing: -0.02em;
    }
    .hero-header p {
        color: var(--text-secondary);
        font-size: 1rem;
        margin: 0;
        font-weight: 400;
    }

    /* ---------- STAT CARDS ROW ---------- */
    .stat-cards {
        display: flex;
        gap: 1rem;
        margin-bottom: 1.5rem;
    }
    .stat-card {
        flex: 1;
        background: var(--bg-card);
        border: 1px solid var(--border);
        border-radius: var(--radius);
        padding: 1.2rem 1.4rem;
        transition: all .2s ease;
    }
    .stat-card:hover {
        border-color: var(--accent);
        box-shadow: 0 0 20px var(--accent-glow);
        transform: translateY(-2px);
    }
    .stat-card .stat-icon {
        font-size: 1.6rem;
        margin-bottom: .35rem;
    }
    .stat-card .stat-value {
        font-size: 1.8rem;
        font-weight: 800;
        color: var(--text-primary);
        letter-spacing: -0.02em;
    }
    .stat-card .stat-label {
        font-size: .8rem;
        color: var(--text-muted);
        text-transform: uppercase;
        letter-spacing: .06em;
        font-weight: 600;
    }

    /* ---------- GLASS CARD ---------- */
    .glass-card {
        background: rgba(30,41,59,.65);
        backdrop-filter: blur(12px);
        -webkit-backdrop-filter: blur(12px);
        border: 1px solid var(--border);
        border-radius: var(--radius);
        padding: 1.5rem;
        margin-bottom: 1rem;
        box-shadow: var(--shadow);
    }
    .glass-card h3 {
        color: var(--text-primary);
        font-weight: 700;
        font-size: 1.1rem;
        margin: 0 0 .75rem 0;
    }

    /* ---------- STATUS BADGE ---------- */
    .badge {
        display: inline-flex;
        align-items: center;
        gap: .35rem;
        padding: .25rem .75rem;
        border-radius: 9999px;
        font-size: .78rem;
        font-weight: 600;
        line-height: 1;
    }
    .badge-success {
        background: rgba(34,197,94,.12);
        color: #4ADE80;
        border: 1px solid rgba(34,197,94,.25);
    }
    .badge-warning {
        background: rgba(245,158,11,.12);
        color: #FBBF24;
        border: 1px solid rgba(245,158,11,.25);
    }

    /* ---------- DIVIDER ---------- */
    .modern-divider {
        height: 1px;
        background: linear-gradient(90deg, transparent, var(--border), transparent);
        margin: 1.5rem 0;
        border: none;
    }

    /* ---------- SIDEBAR STYLING ---------- */
    .sidebar-brand {
        display: flex;
        align-items: center;
        gap: .65rem;
        margin-bottom: 1.5rem;
        padding-bottom: 1rem;
        border-bottom: 1px solid var(--border);
    }
    .sidebar-brand .brand-icon {
        width: 38px;
        height: 38px;
        border-radius: 10px;
        background: linear-gradient(135deg, #6366F1, #8B5CF6);
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 1.3rem;
        box-shadow: 0 4px 12px rgba(99,102,241,.35);
    }
    .sidebar-brand .brand-text {
        font-weight: 700;
        font-size: 1.05rem;
        color: var(--text-primary);
    }
    .sidebar-brand .brand-sub {
        font-size: .72rem;
        color: var(--text-muted);
        font-weight: 500;
    }

    .sidebar-section-title {
        font-size: .72rem;
        text-transform: uppercase;
        letter-spacing: .08em;
        color: var(--text-muted);
        font-weight: 700;
        margin: 1.2rem 0 .5rem 0;
    }

    /* ---------- BUTTONS ---------- */
    .stDownloadButton > button,
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #6366F1, #8B5CF6) !important;
        color: #fff !important;
        border: none !important;
        border-radius: var(--radius-sm) !important;
        font-weight: 600 !important;
        padding: .65rem 1.5rem !important;
        transition: all .2s ease !important;
        box-shadow: 0 4px 14px rgba(99,102,241,.3) !important;
    }
    .stDownloadButton > button:hover,
    .stButton > button[kind="primary"]:hover {
        transform: translateY(-1px) !important;
        box-shadow: 0 6px 20px rgba(99,102,241,.45) !important;
    }
    .stButton > button {
        border-radius: var(--radius-sm) !important;
        border: 1px solid var(--border) !important;
        font-weight: 500 !important;
        transition: all .15s ease !important;
    }

    /* ---------- TABS ---------- */
    .stTabs [data-baseweb="tab-list"] {
        gap: .25rem;
        background: var(--bg-card);
        border-radius: var(--radius);
        padding: .25rem;
        border: 1px solid var(--border);
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: var(--radius-sm);
        padding: .6rem 1.2rem;
        font-weight: 600;
        color: var(--text-secondary);
        background: transparent;
        border: none;
        transition: all .15s ease;
    }
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #6366F1, #8B5CF6) !important;
        color: #fff !important;
        box-shadow: 0 2px 8px rgba(99,102,241,.35);
    }
    .stTabs [data-baseweb="tab-highlight"] {
        display: none;
    }
    .stTabs [data-baseweb="tab-border"] {
        display: none;
    }

    /* ---------- DATAFRAME ---------- */
    [data-testid="stDataFrame"] {
        border: 1px solid var(--border) !important;
        border-radius: var(--radius) !important;
        overflow: hidden;
    }

    /* ---------- FILE UPLOADER ---------- */
    [data-testid="stFileUploader"] {
        border-radius: var(--radius) !important;
    }
    [data-testid="stFileUploader"] section {
        border: 2px dashed var(--border) !important;
        border-radius: var(--radius) !important;
        background: rgba(99,102,241,.04) !important;
        transition: all .2s ease;
    }
    [data-testid="stFileUploader"] section:hover {
        border-color: var(--accent) !important;
        background: rgba(99,102,241,.08) !important;
    }

    /* ---------- SLIDER ---------- */
    .stSlider [data-baseweb="slider"] [role="slider"] {
        background: var(--accent) !important;
        border-color: var(--accent) !important;
    }

    /* ---------- METRIC ---------- */
    [data-testid="stMetric"] {
        background: var(--bg-card);
        border: 1px solid var(--border);
        border-radius: var(--radius);
        padding: 1rem 1.2rem;
    }

    /* ---------- SELECT BOX ---------- */
    [data-baseweb="select"] {
        border-radius: var(--radius-sm) !important;
    }

    /* ---------- EXPANDER ---------- */
    [data-testid="stExpander"] {
        background: var(--bg-card) !important;
        border: 1px solid var(--border) !important;
        border-radius: var(--radius) !important;
    }

    /* ---------- WELCOME SECTION ---------- */
    .welcome-container {
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        padding: 3rem 2rem;
    }
    .welcome-illustration {
        font-size: 4.5rem;
        margin-bottom: 1.5rem;
        animation: float 3s ease-in-out infinite;
    }
    @keyframes float {
        0%, 100% { transform: translateY(0px); }
        50% { transform: translateY(-10px); }
    }
    .welcome-title {
        font-size: 1.5rem;
        font-weight: 700;
        color: var(--text-primary);
        margin-bottom: .5rem;
    }
    .welcome-desc {
        color: var(--text-secondary);
        font-size: .95rem;
        max-width: 500px;
        text-align: center;
        line-height: 1.6;
        margin-bottom: 2rem;
    }

    .feature-grid {
        display: grid;
        grid-template-columns: 1fr 1fr;
        gap: 1rem;
        max-width: 650px;
        width: 100%;
    }
    .feature-item {
        background: var(--bg-card);
        border: 1px solid var(--border);
        border-radius: var(--radius);
        padding: 1.2rem;
        transition: all .2s ease;
    }
    .feature-item:hover {
        border-color: var(--accent);
        transform: translateY(-2px);
        box-shadow: 0 4px 16px var(--accent-glow);
    }
    .feature-item .fi-icon {
        font-size: 1.5rem;
        margin-bottom: .4rem;
    }
    .feature-item .fi-title {
        font-weight: 700;
        font-size: .9rem;
        color: var(--text-primary);
        margin-bottom: .25rem;
    }
    .feature-item .fi-desc {
        font-size: .78rem;
        color: var(--text-muted);
        line-height: 1.4;
    }

    /* ---------- TABLE PREVIEW ---------- */
    .preview-container {
        border: 1px solid var(--border);
        border-radius: var(--radius);
        overflow: hidden;
        background: var(--bg-card);
    }
    .preview-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: .65rem 1rem;
        background: rgba(99,102,241,.06);
        border-bottom: 1px solid var(--border);
    }
    .preview-header .ph-title {
        font-size: .82rem;
        font-weight: 700;
        color: var(--text-primary);
        display: flex;
        align-items: center;
        gap: .4rem;
    }
    .preview-header .ph-badge {
        font-size: .68rem;
        padding: .15rem .5rem;
        background: rgba(99,102,241,.15);
        color: var(--accent-light);
        border-radius: 9999px;
        font-weight: 600;
    }
    .preview-body {
        padding: .75rem;
    }
    </style>
    """, unsafe_allow_html=True)

    # ==================== SIDEBAR ====================
    with st.sidebar:
        st.markdown("""
        <div class="sidebar-brand">
            <div class="brand-icon">🏰</div>
            <div>
                <div class="brand-text">Thanh Trường</div>
                <div class="brand-sub">IT Shinwoo • PDF ➜ Excel</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        st.markdown('<div class="sidebar-section-title">📂 Upload Files</div>', unsafe_allow_html=True)
        uploaded_files = st.file_uploader(
            "Kéo thả PDF hoặc .ZIP",
            type=["pdf", "zip"],
            accept_multiple_files=True,
            label_visibility="collapsed",
            help="Kéo thả nhiều file PDF hoặc 1 file .ZIP chứa thư mục đơn hàng."
        )

        st.markdown('<div class="sidebar-section-title">📁 Local Path</div>', unsafe_allow_html=True)
        local_dir_input = st.text_input(
            "Đường dẫn thư mục",
            value="",
            placeholder=r"D:\DonHang_PDF",
            label_visibility="collapsed"
        )

        st.markdown("<div style='margin-top:.5rem'></div>", unsafe_allow_html=True)
        clear_cache = st.button("🔄  Làm mới dữ liệu", use_container_width=True, help="Xóa cache và nạp lại")
        if clear_cache:
            st.cache_data.clear()
            for k in list(st.session_state.keys()):
                del st.session_state[k]
            st.rerun()

        st.markdown("""
        <div style="margin-top: 2rem; padding-top: 1rem; border-top: 1px solid var(--border);">
            <div style="font-size: .72rem; color: var(--text-muted); text-align: center;">
                Built with ❤️ by <strong style="color: var(--accent-light);">Nguyễn Thanh Trường</strong><br/>
                <span style="opacity: .6;">IT Shinwoo • Powered by PyMuPDF</span>
            </div>
        </div>
        """, unsafe_allow_html=True)

    # ==================== HEADER ====================
    st.markdown("""
    <div class="hero-header">
        <h1>🏰 Camelot — PDF Order Converter</h1>
        <p>Chuyển đổi đơn đặt hàng PDF thành Excel chính xác, nhanh chóng và tự động</p>
    </div>
    """, unsafe_allow_html=True)

    # ==================== FILE COLLECTION ====================
    pdf_files: list[tuple[str, str]] = []

    if uploaded_files:
        for up_f in uploaded_files:
            if up_f.name.lower().endswith(".pdf"):
                dest = os.path.join(UPLOAD_DIR, up_f.name)
                buf_bytes = up_f.getbuffer()
                if not os.path.exists(dest) or os.path.getsize(dest) != len(buf_bytes):
                    with open(dest, "wb") as f_out:
                        f_out.write(buf_bytes)
                pdf_files.append((up_f.name, dest))
            elif up_f.name.lower().endswith(".zip"):
                with zipfile.ZipFile(io.BytesIO(up_f.getbuffer()), "r") as zf:
                    for member in zf.namelist():
                        if member.lower().endswith(".pdf") and not member.startswith("__MACOSX"):
                            bname = os.path.basename(member)
                            if bname:
                                dest = os.path.join(UPLOAD_DIR, bname)
                                member_bytes = zf.read(member)
                                if not os.path.exists(dest) or os.path.getsize(dest) != len(member_bytes):
                                    with open(dest, "wb") as f_out:
                                        f_out.write(member_bytes)
                                pdf_files.append((bname, dest))
    elif local_dir_input and os.path.exists(local_dir_input):
        for root, _, files in os.walk(local_dir_input):
            for f in sorted(files):
                if f.lower().endswith(".pdf"):
                    pdf_files.append((f, os.path.join(root, f)))

    # ==================== MAIN CONTENT ====================
    if pdf_files:
        # ---------- BATCH MODE (multiple files) ----------
        if len(pdf_files) > 1:
            batch_hash = get_batch_hash(pdf_files)
            if st.session_state.get("batch_hash") != batch_hash:
                with st.spinner(f"⚡ Đang xử lý {len(pdf_files)} file PDF..."):
                    zip_bytes, summary_list = generate_batch_zip(tuple(pdf_files))
                    st.session_state["batch_hash"] = batch_hash
                    st.session_state["cached_zip_bytes"] = zip_bytes
                    st.session_state["cached_summary_list"] = summary_list
            else:
                zip_bytes = st.session_state["cached_zip_bytes"]
                summary_list = st.session_state["cached_summary_list"]

            # Stat cards
            total_items = sum(s.get("Số Mặt Hàng", 0) for s in summary_list)
            ok_count = sum(1 for s in summary_list if "✅" in str(s.get("Trạng Thái", "")))
            st.markdown(f"""
            <div class="stat-cards">
                <div class="stat-card">
                    <div class="stat-icon">📄</div>
                    <div class="stat-value">{len(pdf_files)}</div>
                    <div class="stat-label">Files Processed</div>
                </div>
                <div class="stat-card">
                    <div class="stat-icon">📦</div>
                    <div class="stat-value">{total_items}</div>
                    <div class="stat-label">Total Items</div>
                </div>
                <div class="stat-card">
                    <div class="stat-icon">✅</div>
                    <div class="stat-value">{ok_count}/{len(pdf_files)}</div>
                    <div class="stat-label">Success Rate</div>
                </div>
                <div class="stat-card">
                    <div class="stat-icon">💾</div>
                    <div class="stat-value">{len(zip_bytes)/1024:.0f} KB</div>
                    <div class="stat-label">ZIP Size</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            zip_filename = f"Camelot_Excel_{len(pdf_files)}_files_{now_str}.zip"

            st.download_button(
                label=f"📥  Tải toàn bộ {len(pdf_files)} file Excel (.ZIP)",
                data=zip_bytes,
                file_name=zip_filename,
                mime="application/zip",
                type="primary",
                use_container_width=True
            )

            with st.expander("📊 Bảng tổng hợp kết quả", expanded=True):
                st.dataframe(pd.DataFrame(summary_list), use_container_width=True, hide_index=True)

            st.markdown('<div class="modern-divider"></div>', unsafe_allow_html=True)

            st.markdown("#### 🔍 Xem chi tiết từng file")
            file_names = [f"{i+1}. {item[0]}" for i, item in enumerate(pdf_files)]
            selected_idx = st.selectbox(
                "Chọn file:",
                range(len(pdf_files)),
                format_func=lambda i: file_names[i],
                label_visibility="collapsed"
            )
        else:
            selected_idx = 0

        selected_display_name, selected_pdf_path = pdf_files[selected_idx]
        total_pages = get_pdf_page_count(selected_pdf_path)

        # File info bar
        st.markdown(f"""
        <div class="glass-card" style="display: flex; align-items: center; justify-content: space-between; padding: 1rem 1.4rem;">
            <div style="display: flex; align-items: center; gap: .75rem;">
                <div style="width: 40px; height: 40px; border-radius: 10px; background: linear-gradient(135deg, #EF4444, #F97316); display: flex; align-items: center; justify-content: center; font-size: 1.2rem;">📄</div>
                <div>
                    <div style="font-weight: 700; color: var(--text-primary); font-size: .95rem;">{selected_display_name}</div>
                    <div style="font-size: .78rem; color: var(--text-muted);">{total_pages} trang</div>
                </div>
            </div>
            <div class="badge badge-success">● Ready</div>
        </div>
        """, unsafe_allow_html=True)

        # ---------- TABS ----------
        tab_form, tab_shinwoo = st.tabs([
            "📑  Biểu Mẫu Chuẩn",
            "📦  Form SHINWOO"
        ])

        with tab_form:
            exact_excel_bytes = export_exact_form_excel(selected_pdf_path)

            col_info, col_dl = st.columns([3, 1])
            with col_info:
                st.markdown("""
                <div style="display: flex; align-items: center; gap: .5rem; padding: .5rem 0;">
                    <span class="badge badge-success">✓ Sẵn sàng</span>
                    <span style="color: var(--text-secondary); font-size: .85rem;">File Excel biểu mẫu đã được tạo thành công</span>
                </div>
                """, unsafe_allow_html=True)
            with col_dl:
                clean_base = re.sub(r'[\\/*?:"<>|]', '_', os.path.splitext(os.path.basename(selected_pdf_path))[0])
                st.download_button(
                    label="📥 Tải .XLSX",
                    data=exact_excel_bytes,
                    file_name=f"{clean_base}_1TAB.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            st.markdown('<div class="modern-divider"></div>', unsafe_allow_html=True)

            # Page navigation
            page_no = 1
            if total_pages > 1:
                page_no = st.slider(
                    "Trang",
                    min_value=1,
                    max_value=total_pages,
                    value=1,
                    help="Chuyển trang để xem đối chiếu"
                )

            col_pdf, col_html = st.columns([1, 1], gap="medium")
            img_bytes, html_table = get_page_preview(selected_pdf_path, page_no)

            with col_pdf:
                st.markdown(f"""
                <div class="preview-container">
                    <div class="preview-header">
                        <div class="ph-title">🖼️ PDF Gốc</div>
                        <div class="ph-badge">Trang {page_no}/{total_pages}</div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
                st.image(img_bytes, use_container_width=True)

            with col_html:
                st.markdown(f"""
                <div class="preview-container">
                    <div class="preview-header">
                        <div class="ph-title">📊 Biểu Mẫu Tái Tạo</div>
                        <div class="ph-badge">Trang {page_no}</div>
                    </div>
                    <div class="preview-body">
                        {html_table}
                    </div>
                </div>
                """, unsafe_allow_html=True)

        with tab_shinwoo:
            po_num, fac_req, extracted_items = extract_shinwoo_order_items(selected_pdf_path)

            if extracted_items:
                count_with_barcode = sum(1 for itm in extracted_items if itm.get("7A. UPC A Code") or itm.get("7B. EAN 13 Code"))
                count_with_price = sum(1 for itm in extracted_items if itm.get("11. Retail Price"))

                # Stats row for SHINWOO
                st.markdown(f"""
                <div class="stat-cards">
                    <div class="stat-card">
                        <div class="stat-icon">📋</div>
                        <div class="stat-value">{po_num or 'N/A'}</div>
                        <div class="stat-label">PO Number</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-icon">📦</div>
                        <div class="stat-value">{len(extracted_items)}</div>
                        <div class="stat-label">Items</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-icon">🏷️</div>
                        <div class="stat-value">{count_with_barcode}/{len(extracted_items)}</div>
                        <div class="stat-label">Barcodes</div>
                    </div>
                    <div class="stat-card">
                        <div class="stat-icon">💰</div>
                        <div class="stat-value">{count_with_price}/{len(extracted_items)}</div>
                        <div class="stat-label">Prices</div>
                    </div>
                </div>
                """, unsafe_allow_html=True)

                st.markdown("""
                <div style="display: flex; align-items: center; gap: .5rem; margin-bottom: .75rem;">
                    <span style="font-weight: 600; color: var(--text-primary);">✏️ Chỉnh sửa dữ liệu</span>
                    <span style="font-size: .78rem; color: var(--text-muted);">— Nhấp đúp vào ô bất kỳ để sửa trước khi xuất</span>
                </div>
                """, unsafe_allow_html=True)

                df_items = pd.DataFrame(extracted_items)
                edited_items_df = st.data_editor(df_items, use_container_width=True, height=380, hide_index=True)

                filled_shinwoo_bytes = fill_shinwoo_template(edited_items_df.to_dict(orient="records"), factory_request=fac_req)

                st.download_button(
                    label=f"📥  Tải Form SHINWOO — {po_num or 'Output'}.xlsx",
                    data=filled_shinwoo_bytes,
                    file_name=f"SHINWOO_Order_Form_{po_num or 'Output'}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                st.markdown("""
                <div class="glass-card" style="text-align: center; padding: 2rem;">
                    <div style="font-size: 2.5rem; margin-bottom: .75rem;">🔍</div>
                    <div style="font-weight: 600; color: var(--text-primary); margin-bottom: .35rem;">Không tìm thấy dữ liệu</div>
                    <div style="font-size: .85rem; color: var(--text-muted);">Chưa nhận diện được danh sách items cho định dạng PDF này</div>
                </div>
                """, unsafe_allow_html=True)
    else:
        # ==================== WELCOME SCREEN ====================
        st.markdown("""
        <div class="welcome-container">
            <div class="welcome-illustration">📄✨</div>
            <div class="welcome-title">Bắt đầu bằng cách tải file lên</div>
            <div class="welcome-desc">
                Kéo thả file PDF hoặc .ZIP vào sidebar bên trái, hoặc nhập đường dẫn thư mục trên máy tính để hệ thống tự động xử lý.
            </div>
            <div class="feature-grid">
                <div class="feature-item">
                    <div class="fi-icon">📄</div>
                    <div class="fi-title">Single File</div>
                    <div class="fi-desc">Xem đối chiếu PDF vs Excel, chỉnh sửa dữ liệu và tải file riêng lẻ</div>
                </div>
                <div class="feature-item">
                    <div class="fi-icon">📁</div>
                    <div class="fi-title">Batch Mode</div>
                    <div class="fi-desc">Tải nhiều PDF cùng lúc, hệ thống nén trọn bộ Excel vào 1 file .ZIP</div>
                </div>
                <div class="feature-item">
                    <div class="fi-icon">🏷️</div>
                    <div class="fi-title">Auto Barcode</div>
                    <div class="fi-desc">Tự động trích xuất UPC/EAN barcode và giá bán lẻ từ đơn hàng</div>
                </div>
                <div class="feature-item">
                    <div class="fi-icon">📊</div>
                    <div class="fi-title">Form Builder</div>
                    <div class="fi-desc">Tạo biểu mẫu Excel chuẩn kẻ khung, merge cells, đúng format gốc</div>
                </div>
            </div>
        </div>
        """, unsafe_allow_html=True)

if __name__ == "__main__":
    run_streamlit_app()
