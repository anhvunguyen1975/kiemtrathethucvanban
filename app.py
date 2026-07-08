import streamlit as st
import docx
from docx.enum.text import WD_LINE_SPACING
from docx import Document
from docx.shared import Cm, Pt, Inches, RGBColor
from docx.oxml.ns import qn  
from docx.shared import Pt, RGBColor
from docx.oxml import OxmlElement

import re
import io
import os
import unicodedata  
import json  
#from reportlab.pdfbase import pdfmetrics
#from reportlab.pdfbase.ttfonts import TTFont
#from reportlab.pdfgen import canvas
#from reportlab.lib.pagesizes import letter

st.set_page_config(page_title="Kiểm tra thể thức văn bản NĐ 30", page_icon="💧", layout="wide", initial_sidebar_state="expanded")

# --- CẤU HÌNH GIAO DIỆN STREAMLIT ---
custom_css = """
<style>
  .stAppDeployButton {display: none !important;}
  #MainMenu { visibility: hidden !important; }
  header {visibility: hidden;}
  footer {visibility: hidden;}
  .viewerBadge_container__171of {display: none !important;}
  
  [data-testid="collapsedControl"] {
      display: flex !important;
      visibility: visible !important;
      z-index: 999999 !important; 
      background-color: transparent !important;
  }
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

chuc_vu_keywords = ["GIÁM ĐỐC", "PHÓ GIÁM ĐỐC", "CHỦ TỊCH", "TRƯỞNG PHÒNG", "TỔNG GIÁM ĐỐC", "PHÓ TỔNG GIÁM ĐỐC"]
# --- DANH MỤC LOẠI VĂN BẢN CHUẨN NĐ 30 ---
LOAI_VAN_BAN_CHUAN = {
    "NQ": "Nghị quyết", "QĐ": "Quyết định", "CT": "Chỉ thị",
    "QC": "Quy chế", "QyĐ": "Quy định", "TB": "Thông báo", "HD": "Hướng dẫn", 
    "CTr": "Chương trình", "CTY": "Công ty", "KH": "Kế hoạch", "PA": "Phương án", 
    "BC": "Báo cáo", "BB": "Biên bản", "TTr": "Tờ trình", "HĐ": "Hợp đồng", 
    "GUQ": "Giấy ủy quyền", "GM": "Giấy mời", "GGT": "Giấy giới thiệu"
}
MAP_CHUAN_HOA_LOAI = {k.upper(): k for k in LOAI_VAN_BAN_CHUAN.keys()}
MAP_CHUAN_HOA_LOAI.update({"TTR": "TTr", "CTR": "CTr", "QYĐ": "QyĐ"})

CONFIG_FILE = "bawa_config.json"

def load_hidden_config():
    if not os.path.exists(CONFIG_FILE):
        default_config = {
            "ten_co_quan_me": "CÔNG TY CỔ PHẦN CẤP NƯỚC BẠC LIÊU",
            "dia_danh_chuan": "Bạc Liêu",
            "danh_sach_phong_ban": [
                {"ten": "Phòng Tổ chức Hành chính", "viet_tat": "TCHC"},
                {"ten": "Phòng Kế toán", "viet_tat": "KT"},
                {"ten": "Phòng Kinh doanh", "viet_tat": "KD"},
                {"ten": "Phòng Kế hoạch Kỹ thuật", "viet_tat": "KHKT"},
                {"ten": "Ban kiểm soát", "viet_tat": "BKS"},
                {"ten": "Xí nghiệp Cấp nước", "viet_tat": "XNCN"}
            ]
        }
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(default_config, f, ensure_ascii=False, indent=4)
        return default_config
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

HIDDEN_CONFIG = load_hidden_config()

def init_vietnamese_fonts():
    win_font_reg, win_font_bold = "C:\\Windows\\Fonts\\arial.ttf", "C:\\Windows\\Fonts\\arialbd.ttf"
    if os.path.exists(win_font_reg) and os.path.exists(win_font_bold):
        pdfmetrics.registerFont(TTFont('Arial', win_font_reg))
        pdfmetrics.registerFont(TTFont('Arial-Bold', win_font_bold))
        return 'Arial', 'Arial-Bold'
    return 'Helvetica', 'Helvetica-Bold'

PDF_FONT_REG, PDF_FONT_BOLD = init_vietnamese_fonts()

def set_font_times(run):
    run.font.name = 'Times New Roman'
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn('w:ascii'), 'Times New Roman')
    rFonts.set(qn('w:hAnsi'), 'Times New Roman')
    rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    rFonts.set(qn('w:cs'), 'Times New Roman')

# --- HÀM KIỂM TRA SÂU ---
def analyze_document_v6(doc):
    success_items = []
    error_list = []
    warning_list = []
    ambiguous_dict = {} 

    for idx, section in enumerate(doc.sections):
        top_cm = round(section.top_margin.cm, 2) if section.top_margin else 0
        bottom_cm = round(section.bottom_margin.cm, 2) if section.bottom_margin else 0
        left_cm = round(section.left_margin.cm, 2) if section.left_margin else 0
        right_cm = round(section.right_margin.cm, 2) if section.right_margin else 0
        
        margin_errors = []
        if not (2.0 <= top_cm <= 2.5): margin_errors.append(f"Lề trên: {top_cm}cm")
        if not (2.0 <= bottom_cm <= 2.5): margin_errors.append(f"Lề dưới: {bottom_cm}cm")
        if not (3.0 <= left_cm <= 3.5): margin_errors.append(f"Lề trái: {left_cm}cm")
        if not (1.5 <= right_cm <= 2.0): margin_errors.append(f"Lề phải: {right_cm}cm")
        
        if margin_errors:
            error_list.append(f"❌ **Sai căn lề trang:** " + " | ".join(margin_errors))
        else:
            success_items.append(f"Kích thước căn lề trang Section {idx+1} (Đạt chuẩn)")

    wrong_fonts_detected = set()
    all_paragraphs = []
    for p in doc.paragraphs:
        if p.text.strip(): all_paragraphs.append(p)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    if paragraph.text.strip(): all_paragraphs.append(paragraph)

    for p in all_paragraphs:
        for run in p.runs:
            if run.text.strip() and run.font.name and run.font.name != "Times New Roman":
                wrong_fonts_detected.add(run.font.name)

    if wrong_fonts_detected:
        error_list.append(f"❌ **Sai Phông chữ:** Phát hiện phông lạ {list(wrong_fonts_detected)} → Bắt buộc dùng **Times New Roman**.")
    else:
        success_items.append("Toàn bộ phông chữ bản thảo (Đạt chuẩn: Times New Roman)")

    def is_paragraph_bold(p):
        style_bold = False
        try:
            if p.style and p.style.font and p.style.font.bold: style_bold = True
        except: pass
        text_runs = [run for run in p.runs if run.text.strip()]
        if not text_runs: return False
        for run in text_runs:
            rb = run.bold if run.bold is not None else style_bold
            if not rb: return False
        return True

    doc_text = "\n".join([p.text for p in all_paragraphs])
    checked_departments = set()
    checked_departments_line = set()
    has_noi_nhan = False
    has_nguoi_ky = False
    
    for idx, p in enumerate(all_paragraphs):
        line_clean = p.text.replace("|", "").strip()
        if not line_clean: continue
        line_clean_no_underscore = line_clean.strip("_").strip()
        line_upper = unicodedata.normalize('NFC', line_clean_no_underscore.upper())
        line_clean_spaces = re.sub(r'\s+', ' ', line_clean_no_underscore)
        
        # V6.6: Kiểm tra đường kẻ ngang dưới tên Cơ quan theo NĐ 30
        is_agency_tail = False
        valid_tails_lower = ["cấp nước bạc liêu", "công ty cổ phần cấp nước bạc liêu"] + [pb["ten"].lower() for pb in HIDDEN_CONFIG["danh_sach_phong_ban"]]
        if any(line_clean_no_underscore.lower().endswith(tail) for tail in valid_tails_lower):
                is_agency_tail = True
                
        if is_agency_tail and line_clean_no_underscore.lower() not in checked_departments_line:
                has_line = False
                # Kiểm tra xem có dấu gạch dưới hoặc Shape XML hay không
                if "_" in p.text: has_line = True
                if "<w:drawing" in p._p.xml or "<v:line" in p._p.xml: has_line = True
                if "<w:pBdr" in p._p.xml: has_line = True
                try:
                    if p._cell and "<w:tcBorders" in p._cell._tc.xml: has_line = True 
                except: pass     
                # Quét thêm đoạn văn ngay bên dưới (vì dòng kẻ thường rớt xuống dòng kế tiếp)
                if not has_line and idx + 1 < len(all_paragraphs):
                    next_p = all_paragraphs[idx+1]
                    if "_" in next_p.text or "<w:drawing" in next_p._p.xml or "<v:line" in next_p._p.xml:
                        has_line = True
                if not has_line:
                    #warning_list.append(f"⚠️ **Thiếu đường kẻ ngang** phía dưới tên cơ quan ban hành văn bản `[{line_clean_no_underscore}]`. **<span style='color: red;'>Vui lòng mở file thêm thủ công</span>** để đúng chuẩn NĐ 30.")
                    checked_departments_line.add(line_clean_no_underscore.lower())
                
        # Kiểm tra Quốc hiệu, Tiêu ngữ, Tên công ty
        if "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM" in line_upper:
            if not is_paragraph_bold(p): error_list.append(f"❌ Sai Quốc hiệu: `[{line_clean_no_underscore}]` → Bắt buộc phải **IN ĐẬM**.")
        if "độc lập" in line_clean_no_underscore.lower() and "hạnh phúc" in line_clean_no_underscore.lower():
            if not is_paragraph_bold(p): error_list.append(f"❌ Sai Tiêu ngữ: `[{line_clean_no_underscore}]` → Bắt buộc phải **IN ĐẬM**.")
        if line_clean_no_underscore.lower() in ["công ty cổ phần", "cấp nước bạc liêu", "công ty cổ phần cấp nước bạc liêu"]:
            if line_clean_no_underscore.lower() not in checked_departments:
                if line_clean_no_underscore != line_clean_no_underscore.upper() or not is_paragraph_bold(p):
                    error_list.append(f"❌ Sai định dạng tên Công ty: `[{line_clean_no_underscore}]` → Phải viết **HOA TOÀN BỘ** và **IN ĐẬM**.")
                checked_departments.add(line_clean_no_underscore.lower())

        # Kiểm tra chức vụ Người ký văn bản
        if any(kv in line_clean_spaces for kv in chuc_vu_keywords):
            if "KÍNH GỬI" not in line_upper:
                has_nguoi_ky = True

        # Kiểm tra quy định KÍNH GỬI (NĐ 30)
        if line_clean_no_underscore.lower().startswith("kính gửi"):
            if "KÍNH GỬI" in line_clean_no_underscore:
                error_list.append(f"❌ Sai định dạng Kính gửi: `[{line_clean_no_underscore[:20]}...]` → Theo NĐ 30, chữ 'Kính gửi' phải in thường, không viết HOA toàn bộ.")
            if is_paragraph_bold(p):
                error_list.append(f"❌ Sai kiểu chữ Kính gửi: `[{line_clean_no_underscore[:20]}...]` → Bắt buộc dùng kiểu chữ thường, đứng (không in đậm).")

        # Kiểm tra Nơi nhận
        if line_clean_no_underscore.lower().startswith("nơi nhận"):
            has_noi_nhan = True
            is_italic = any(r.italic for r in p.runs if r.text.strip())
            is_bold = is_paragraph_bold(p)
            if not (is_bold and is_italic):
                error_list.append(f"❌ Khối `Nơi nhận:` chưa đúng định dạng (Yêu cầu: Vừa **In đậm** vừa *In nghiêng*).")

        # Kiểm tra Số hiệu
        if re.match(r"^\s*Số\s*:", line_clean_no_underscore, re.IGNORECASE):
            valid_vts = [pb["viet_tat"].upper() for pb in HIDDEN_CONFIG["danh_sach_phong_ban"]]
            notation_match = re.search(r"/\s*([A-ZĐa-zđ0-9]+)(?:\s*-\s*([A-ZĐa-zđ0-9]*))?", line_clean_no_underscore)
            if notation_match:
                full_notation, agency_raw, dept_raw = notation_match.group(0), notation_match.group(1), notation_match.group(2)
                agency_upper, dept = agency_raw.upper(), dept_raw.upper() if dept_raw else ""
                is_notation_err = False
                
                if agency_upper not in MAP_CHUAN_HOA_LOAI:
                    error_list.append(f"❌ Số hiệu sai loại văn bản: `[{agency_raw}]`")
                    is_notation_err = True

                if dept == "":
                    is_notation_err = True
                    correct_agency = MAP_CHUAN_HOA_LOAI.get(agency_upper, agency_raw)
                    choices = [f"Chỉ {correct_agency} (Bỏ dấu gạch ngang)"] + [f"{pb['ten']} ({pb['viet_tat']})" for pb in HIDDEN_CONFIG["danh_sach_phong_ban"]]
                    ambiguous_dict[full_notation] = choices
                    error_list.append(f"⚠️ Số hiệu trống phòng ban: `[{full_notation}]`")
                else:
                    if dept not in valid_vts:
                        is_notation_err = True
                        error_list.append(f"❌ Số hiệu không khớp danh mục BAWACO: `[{full_notation}]`")
                if not is_notation_err:
                    success_items.append(f"Số ký hiệu văn bản chính thức ({full_notation} - Đạt chuẩn)")

    if "hành chánh" in doc_text.lower():
        error_list.append("❌ Sai từ ngữ hành chính: phát hiện từ **hành chánh**.")
        
    if not has_noi_nhan:
        warning_list.append("⚠️ Chưa tìm thấy khối danh mục tiêu đề `Nơi nhận:` (Hệ thống sẽ tự động tạo khi bấm nút Autofix).")

    if not has_nguoi_ky:
        warning_list.append("⚠️ **Cảnh báo:** Chưa phát hiện tiêu đề chức vụ của **Người ký** văn bản ở cuối trang (Ví dụ: GIÁM ĐỐC, TỔNG GIÁM ĐỐC...).")
    else:
        success_items.append("Chức vụ Người ký văn bản (Đã ghi nhận)")

    
    return success_items, error_list, warning_list, ambiguous_dict

# --- GIAO DIỆN CHÍNH ---
st.title("KIỂM TRA THỂ THỨC VĂN BẢN (NĐ 30)")
st.markdown("🚀 **Phiên bản V6.6:** Bổ sung cơ chế quét và tạo **đường kẻ ngang (1/3 - 1/2 độ dài)** dưới tên cơ quan ban hành.")
st.markdown("---")

with st.sidebar:
    st.markdown("### 📁 BỘ NẠP VĂN BẢN")
    uploaded_file = st.file_uploader("Chọn file bản thảo văn bản (.docx):", type=["docx"])
    st.markdown("---")
    st.markdown("### 🏢 DANH MỤC BAWACO")
    with st.expander("Xem bảng viết tắt chuẩn", expanded=False):
        st.caption(f"**Công ty:** {HIDDEN_CONFIG['ten_co_quan_me']}")
        for pb in HIDDEN_CONFIG["danh_sach_phong_ban"]:
            st.caption(f"• **{pb['viet_tat']}**: {pb['ten']}")

user_resolutions = {}

if uploaded_file is not None:
    doc = Document(uploaded_file)
    success_items, error_list, warning_list, ambiguous_dict = analyze_document_v6(doc)
    
    st.markdown("### 📊 Kết quả Phân tích & Kiểm tra toàn diện")
    col_overview, col_highlight = st.columns([1, 1])
    
    with col_overview:
        st.write("**📂 Đang xử lý file:** `" + uploaded_file.name + "`")
        with st.expander("Các tiêu chí đã đạt chuẩn", expanded=True):
            for item in success_items: st.write(f"✔️ {item}", unsafe_allow_html=True)
            
    with col_highlight:
        st.markdown("### 🖍️ DANH SÁCH LỖI / SAI LỆCH THỂ THỨC")
        if error_list or warning_list:
            with st.chat_message("assistant", avatar="📝"):
                for item in error_list: st.markdown(item, unsafe_allow_html=True)
                for item in warning_list: st.markdown(item, unsafe_allow_html=True)
        else:
            st.success("✨ Xuất sắc! Văn bản đạt điểm tối đa!")
        
        if ambiguous_dict:
            for notation, choices in ambiguous_dict.items():
                user_resolutions[notation] = st.selectbox(f"Chọn đuôi đúng cho '{notation}':", options=choices)

    # --- HỘP CÔNG CỤ XỬ LÝ (AUTOFIX V6.6) ---
    st.markdown("### 🛠️ HỘP CÔNG CỤ XỬ LÝ TỰ ĐỘNG CHUẨN HOÁ")
    if st.button("🪄 TỰ ĐỘNG FIX TOÀN DIỆN (NĐ30)", type="primary"):
        with st.spinner("Đang định dạng lại lề trang, phông chữ và nội dung..."):
            
            # --- BƯỚC 1: CĂN LỀ TRANG CHUẨN ---
            for section in doc.sections:
                section.top_margin = Cm(2.0)
                section.bottom_margin = Cm(2.0)
                section.left_margin = Cm(3.0)
                section.right_margin = Cm(1.5)

            # --- BƯỚC 2: AUTOFIX THIẾU KHỐI "NƠI NHẬN:" ---
            has_noi_nhan_local = False
            for p in doc.paragraphs:
                if p.text.strip().lower().startswith("nơi nhận:"):
                    has_noi_nhan_local = True
                    break
            
            if not has_noi_nhan_local:
                for p in doc.paragraphs:
                    text_strip = p.text.strip()
                    if text_strip.startswith("-") or text_strip.startswith("•") or "lưu:" in text_strip.lower():
                        new_p = p.insert_paragraph_before()
                        r = new_p.add_run("Nơi nhận:")
                        r.bold = True
                        r.italic = True
                        r.font.size = Pt(12)
                        set_font_times(r)
                        break

            # --- BƯỚC 3: QUÉT HEADER & CHÍNH TẢ ---
            
            def fix_para(p, paragraph_index):
                text_clean = p.text.replace("|", "").strip()
                if not text_clean: return
                # Lọc bỏ các dấu gạch dưới cũ để code không bị nhầm lẫn khi format lại
                text_clean_no_underscore = text_clean.strip("_").strip()
                if not text_clean_no_underscore: return
                
                text_upper = unicodedata.normalize('NFC', text_clean_no_underscore.upper())
                text_lower = text_clean_no_underscore.lower()
                text_clean_spaces = re.sub(r'\s+', ' ', text_clean_no_underscore)
                orig_size = p.runs[0].font.size if p.runs and p.runs[0].font.size else Pt(13)

                # ÉP CHUẨN "KÍNH GỬI"
                if text_lower.startswith("kính gửi") or text_lower.startswith("kính gởi"):
                    for r in p.runs:
                        r.bold = False
                        r.italic = False
                        r.font.size = Pt(14)
                        set_font_times(r)
                        temp_text = r.text
                        for variant in ["KÍNH GỬI", "Kính Gửi", "KÍNH GỞI", "Kính Gởi"]:
                            temp_text = temp_text.replace(variant, "Kính gửi")
                        r.text = temp_text
                        
                    p.alignment = 0 
                    p.paragraph_format.left_indent = None 
                    p.paragraph_format.first_line_indent = Cm(1.27) 
                    return

                # V6.6: ÉP CHUẨN ĐƯỜNG KẺ NGANG DƯỚI TÊN CƠ QUAN
                if paragraph_index < 2:
                    if any(x in text_upper for x in ["CÔNG TY", "CẤP NƯỚC", "PHÒNG", "BAN", "XÍ NGHIỆP", "TRUNG TÂM"]):
                        if not any(x in text_upper for x in ["CỘNG HÒA", "ĐỘC LẬP", "SỐ:", "NGÀY", "THÁNG", "NĂM", "CĂN CỨ", "CHỦ TỊCH", "GIÁM ĐỐC", "BAN QUẢN LÝ", "KÍNH GỬI", "V/V"]):
                            hoa_text = text_clean_no_underscore.upper()
                            
                            # Xác định xem dòng này có phải là dòng cuối của khối cơ quan không
                            needs_line = False
                            valid_tails = ["CẤP NƯỚC BẠC LIÊU", "CÔNG TY CỔ PHẦN CẤP NƯỚC BẠC LIÊU"] + [pb["ten"].upper() for pb in HIDDEN_CONFIG["danh_sach_phong_ban"]]
                            if any(hoa_text.endswith(tail) for tail in valid_tails):
                                needs_line = True
                                p.text = "" 
                                parts = hoa_text.split('\n')
                                for i, part in enumerate(parts):
                                    r = p.add_run(part)
                                    r.bold = True
                                    r.font.size = Pt(13)
                                    set_font_times(r)
                                
                                    if i < len(parts) - 1:
                                        p.add_run('\n')
                                    
                                    p.alignment = 1 
                                # 3. Thêm đường kẻ bằng cách ngắt dòng trong cùng 1 paragraph
                                                
                        return 

                if "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM" in text_upper or re.search(r"CỘNG\s*H[ÒOÀA]+\s*XÃ\s*HỘI", text_upper):
                    p.text = ""
                    r = p.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
                    r.bold = True; r.font.size = orig_size; set_font_times(r)
                    p.alignment = 1
                    return
                        
                if "độc lập" in text_lower and "hạnh phúc" in text_lower:
                    p.text = ""
                    r = p.add_run("Độc lập - Tự do - Hạnh phúc")
                    r.bold = True; r.font.size = Pt(14); set_font_times(r)
                    rPr = r._r.get_or_add_rPr()
                    u = OxmlElement('w:u')
                    u.set(qn('w:val'), 'single') 
                    u.set(qn('w:sz'), '0.1')
                    u.set(qn('w:space'), '12')  
                    rPr.append(u)
                    p.alignment = 1
                    return

                if "ngày" in text_lower and "tháng" in text_lower and "năm" in text_lower:
                    if len(text_clean_no_underscore) < 70 and not any(x in text_lower for x in ["căn cứ", "luật", "nghị định", "quyết định", "thông tư", "v/v", "về việc"]):
                        p.text = "Cà Mau, ngày    tháng    năm 2026"
                        for r in p.runs: r.italic = True; r.font.size = Pt(13); set_font_times(r)
                        p.alignment = 1
                        return

                if re.match(r"^\s*Số\s*:", text_clean_no_underscore, re.IGNORECASE) or "Số:" in text_clean_no_underscore:
                    temp_text = p.text
                    def fix_notation(match):
                        raw, ag_raw, dt_raw = match.group(0), match.group(1), match.group(2)
                        correct_agency = MAP_CHUAN_HOA_LOAI.get(ag_raw.upper(), ag_raw)
                        valid_vts = [pb["viet_tat"] for pb in HIDDEN_CONFIG["danh_sach_phong_ban"]]
                        valid_vts_upper = [v.upper() for v in valid_vts]
                        try:
                            if raw in user_resolutions:
                                sel = user_resolutions[raw]
                                if "Bỏ dấu gạch ngang" in sel: return f"/{correct_agency}"
                                for pb in HIDDEN_CONFIG["danh_sach_phong_ban"]:
                                    if f"({pb['viet_tat']})" in sel: return f"/{correct_agency}-{pb['viet_tat']}"
                        except NameError: pass
                        dt = dt_raw.upper() if dt_raw else ""
                        if dt in valid_vts_upper: return f"/{correct_agency}-{valid_vts[valid_vts_upper.index(dt)]}"
                        return f"/{correct_agency}-{dt}" if dt else f"/{correct_agency}"

                    temp_text = re.sub(r"/\s*([A-ZĐa-zđ0-9]+)(?:\s*-\s*([A-ZĐa-zđ0-9]*))?", fix_notation, temp_text)
                    if temp_text != p.text:
                        was_bold = any(r.bold for r in p.runs)
                        p.text = temp_text
                        for r in p.runs: r.bold = was_bold; r.font.size = orig_size; set_font_times(r)
                    return

                for r in p.runs:
                    if not r.text.strip(): continue
                    if re.search(r"\bhành\s+chánh\b", r.text, flags=re.IGNORECASE):
                        r.text = re.sub(r"\bhành\s+chánh\b", "hành chính", r.text, flags=re.IGNORECASE)
                    for c in ["Bạc Liêu", "Cà Mau", "Hà Nội", "Hồ Chí Minh"]:
                        if c.upper() not in r.text:
                            r.text = re.compile(rf"\b{re.escape(c)}\b", re.IGNORECASE).sub(c, r.text)
                    set_font_times(r)


            for idx, p in enumerate(doc.paragraphs): fix_para(p, idx)
            for t in doc.tables:
                for row in t.rows:
                    for cell in row.cells:
                        for idx, p in enumerate(cell.paragraphs):
                            fix_para(p, idx)

            # --- BƯỚC 4: POST-PROCESSING TÁCH BIỆT ---
            for p in doc.paragraphs:
                text_clean = p.text.strip()
                text_lower = text_clean.lower()
                if len(text_clean) > 0 and (text_lower.startswith("nơi nhận:") or text_lower.startswith("nơi nhận :")):
                    p.text = "" 
                    r = p.add_run("Nơi nhận:")
                    r.bold = True
                    r.italic = True
                    r.font.size = Pt(12)
                    set_font_times(r)
                    p.alignment = 0 
                elif text_clean.startswith("-") or text_clean.startswith("•") or "lưu:" in text_lower:
                    if len(text_clean) < 80: 
                        for r in p.runs:
                            r.font.size = Pt(11) 
                            set_font_times(r)

            # --- Cập nhật Autofix chữ ký (Phiên bản ổn định) ---
            # Tạo danh sách các đoạn văn cần kiểm tra từ chính đối tượng doc
            elements_to_check = list(doc.paragraphs)
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        elements_to_check.extend(cell.paragraphs)

            # Quét qua tất cả các đoạn văn
            for p in elements_to_check:
                text_clean = p.text.replace("|", "").strip()
                if not text_clean: continue
                
                # Kiểm tra xem có chứa chức vụ nào trong list không
                if any(kv in text_clean.upper() for kv in chuc_vu_keywords):
                    # Loại trừ trường hợp là "KÍNH GỬI"
                    if "KÍNH GỬI" not in text_clean.upper(): 
                        for r in p.runs: 
                            r.bold = True # In đậm chức vụ
                            # Thêm định dạng font nếu cần
                            r.font.name = 'Times New Roman' 
                        
                        # Căn giữa nếu đoạn văn đó không phải là Nơi nhận
                        if "NƠI NHẬN" not in text_clean.upper():
                            p.alignment = 1 # Căn giữa (Center)

            output = io.BytesIO()
            doc.save(output)
            st.success("✅ **Tuyệt vời!** Văn bản đã được dọn dẹp sạch sẽ theo đúng NĐ 30/2020/NĐ-CP (Kể cả đường kẻ ngang của Cơ quan)!")
            
            st.download_button(
                label="⬇️ TẢI XUỐNG VĂN BẢN CHUẨN (DOCX)",
                data=output.getvalue(),
                file_name="Da_Chuan_Hoa_" + uploaded_file.name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
else:
    st.info("👋 Xin chào! Hãy tải lên một bản thảo văn bản để hệ thống bắt đầu kiểm tra lỗi thể thức!")
