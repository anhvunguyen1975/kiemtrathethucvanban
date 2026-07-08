import streamlit as st
from docx import Document
from docx.shared import Cm, Pt
from docx.oxml.ns import qn  # Cấu hình XML Font
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
import re
import io
import os
import unicodedata  
import json  

# Thêm dòng này ở ngay đầu ứng dụng
st.set_page_config(
    page_title="Kiểm tra thể thức văn bản NĐ 30",
    page_icon="💧", # Đã đổi thành icon giọt nước
    layout="centered"
)
# --- CÁC IMPORT THƯ VIỆN REPORTLAB ---
# from reportlab.pdfbase import pdfmetrics
# from reportlab.pdfbase.ttfonts import TTFont
# from reportlab.pdfgen import canvas
# from reportlab.lib.pagesizes import letter

# Ẩn menu, footer, logo Streamlit và thanh trang trí
hide_st_style = """
<style>
#MainMenu {visibility: hidden;}
#GithubIcon {visibility: hidden;}
.styles_viewerBadge__1yB5_, .viewerBadge_link__1S137, .viewerBadge_text__1JaDK {display: none !important;}
footer {visibility: hidden;}
header {visibility: hidden;}
#stDecoration {display: none;}
</style>
"""
st.markdown(hide_st_style, unsafe_allow_html=True)

# --- DANH MỤC LOẠI VĂN BẢN CHUẨN NĐ 30 ---
LOAI_VAN_BAN_CHUAN = {
    "NQ": "Nghị quyết (cá biệt)", "QĐ": "Quyết định (cá biệt)", "CT": "Chỉ thị",
    "QC": "Quy chế", "QyĐ": "Quy định", "QYĐ": "Quy định", "TC": "Thông cáo",
    "TB": "Thông báo", "HD": "Hướng dẫn", "CTr": "Chương trình", "CTR": "Chương trình",
    "CTY": "Công ty", "KH": "Kế hoạch", "PA": "Phương án", "ĐA": "Đề án", "DA": "Dự án",
    "BC": "Báo cáo", "BB": "Biên bản", "TTr": "Tờ trình", "TTR": "Tờ trình",
    "HĐ": "Hợp đồng", "CĐ": "Công điện", "BGN": "Bản ghi nhớ", "BTT": "Bản thỏa thuận",
    "GUQ": "Giấy ủy quyền", "GM": "Giấy mời", "GGT": "Giấy giới thiệu",
    "GNP": "Giấy nghỉ phép", "PG": "Phiếu gửi", "PC": "Phiếu chuyển", "PB": "Phiếu báo"
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
                {"ten": "Phòng Quản lý Mạng lưới Cấp nước", "viet_tat": "QLML CN"},
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

def clean_text_for_pdf(text):
    text = re.sub(r'<[^>]*>', '', text)
    for emo in ["💡", "🎉", "⚠️", "🛡️", "📂", "✔️", "✨", "📝", "❌"]:
        text = text.replace(emo, "")
    return text.strip()

# --- HÀM KIỂM TRA SÂU (BỔ SUNG CĂN LỀ & FONT) ---
def analyze_document_v5(doc):
    success_items = []
    error_list = []
    warning_list = []
    ambiguous_dict = {} 

    # 1. KIỂM TRA CĂN LỀ TRANG (MARGINS)
    for idx, section in enumerate(doc.sections):
        # Chuyển đổi từ đơn vị inch/Twips sang cm để dễ so sánh
        top_cm = round(section.top_margin.cm, 2) if section.top_margin else 0
        bottom_cm = round(section.bottom_margin.cm, 2) if section.bottom_margin else 0
        left_cm = round(section.left_margin.cm, 2) if section.left_margin else 0
        right_cm = round(section.right_margin.cm, 2) if section.right_margin else 0
        
        margin_errors = []
        if not (2.0 <= top_cm <= 2.5): margin_errors.append(f"Lề trên: {top_cm}cm (Chuẩn: 2.0-2.5cm)")
        if not (2.0 <= bottom_cm <= 2.5): margin_errors.append(f"Lề dưới: {bottom_cm}cm (Chuẩn: 2.0-2.5cm)")
        if not (3.0 <= left_cm <= 3.5): margin_errors.append(f"Lề trái: {left_cm}cm (Chuẩn: 3.0-3.5cm)")
        if not (1.5 <= right_cm <= 2.0): margin_errors.append(f"Lề phải: {right_cm}cm (Chuẩn: 1.5-2.0cm)")
        
        if margin_errors:
            error_list.append(f"❌ **Sai kích thước căn lề trang (Section {idx+1}):** " + " | ".join(margin_errors) + " → Cần tự động căn chỉnh lại.")
        else:
            success_items.append(f"Kích thước căn lề trang Section {idx+1} (Trên:{top_cm}cm, Dưới:{bottom_cm}cm, Trái:{left_cm}cm, Phải:{right_cm}cm - Đạt chuẩn NĐ 30)")

    # 2. KIỂM TRA PHÔNG CHỮ TOÀN VĂN BẢN (FONTS)
    wrong_fonts_detected = set()
    all_paragraphs = []
    
    # Gom paragraph chính
    for p in doc.paragraphs:
        if p.text.strip(): all_paragraphs.append(p)
    # Gom paragraph trong bảng
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
        error_list.append(f"❌ **Sai Phông chữ:** Phát hiện phông lạ {list(wrong_fonts_detected)} trong bản thảo → Nghị định 30 bắt buộc dùng **Times New Roman**.")
    else:
        success_items.append("Toàn bộ phông chữ bản thảo (Đạt chuẩn: Times New Roman)")

    # 3. CÁC LOGIC KIỂM TRA THỂ THỨC SẴN CÓ CỦA V4.9
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

    quoc_hieu_perfect = quoc_hieu_error = tieu_ngu_perfect = tieu_ngu_error = False
    doc_text = "\n".join([p.text for p in all_paragraphs])
    checked_departments = set()

    for p in all_paragraphs:
        line_clean = p.text.replace("|", "").strip()
        if not line_clean: continue
        line_upper = unicodedata.normalize('NFC', line_clean.upper())
        
        if "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM" in line_upper or "CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM" in line_upper or re.search(r"CỘNG\s*H[ÒÓỎÕỌOÔỒỐỔỖỘƠỜỚỞỠỢ]*[AÀÁẢÃẠ]*\s*XÃ\s*HỘI", line_upper):
            if "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM" in line_clean or "CỘNG HOÀ XÃ HỘI CHỦ NGHĨA VIỆT NAM" in line_clean:
                if is_paragraph_bold(p): quoc_hieu_perfect = True
                else: quoc_hieu_error = True; error_list.append(f"❌ Sai Quốc hiệu: `[{line_clean}]` → Bắt buộc phải **IN ĐẬM**.")
            else:
                quoc_hieu_error = True; error_list.append(f"❌ Sai chính tả Quốc hiệu: `[{line_clean}]` → Sửa thành: **CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM**")

        if "độc lập" in line_clean.lower() and "hạnh phúc" in line_clean.lower():
            if re.search(r"Độc\s*lập\s*[\-\–\—]\s*Tự\s*do\s*[\-\–\—]\s*Hạnh\s*phúc", line_clean):
                if is_paragraph_bold(p): tieu_ngu_perfect = True
                else: tieu_ngu_error = True; error_list.append(f"❌ Sai Tiêu ngữ: `[{line_clean}]` → Bắt buộc phải **IN ĐẬM**.")
            else:
                tieu_ngu_error = True; error_list.append(f"❌ Sai cấu trúc Tiêu ngữ: `[{line_clean}]` → Sửa thành: **Độc lập - Tự do - Hạnh phúc**")

        if line_clean.lower() in ["công ty cổ phần", "cấp nước bạc liêu", "công ty cổ phần cấp nước bạc liêu"]:
            if line_clean.lower() not in checked_departments:
                if line_clean != line_clean.upper() or not is_paragraph_bold(p):
                    error_list.append(f"❌ Sai định dạng tên Công ty: `[{line_clean}]` → Phải viết **HOA TOÀN BỘ** và **IN ĐẬM**.")
                checked_departments.add(line_clean.lower())

        for pb in HIDDEN_CONFIG["danh_sach_phong_ban"]:
            if line_clean.lower() == pb["ten"].lower() and pb["ten"] not in checked_departments:
                if line_clean != pb["ten"].upper() or not is_paragraph_bold(p):
                    error_list.append(f"❌ Sai định dạng đơn vị: `[{line_clean}]` → Phải viết **HOA TOÀN BỘ** và **IN ĐẬM**.")
                checked_departments.add(pb["ten"])

        if re.match(r"^\s*Số\s*:", line_clean, re.IGNORECASE):
            valid_vts = [pb["viet_tat"].upper() for pb in HIDDEN_CONFIG["danh_sach_phong_ban"]]
            notation_match = re.search(r"/\s*([A-ZĐa-zđ0-9]+)(?:\s*-\s*([A-ZĐa-zđ0-9]*))?", line_clean)
            if notation_match:
                full_notation, agency_raw, dept_raw = notation_match.group(0), notation_match.group(1), notation_match.group(2)
                agency_upper, dept = agency_raw.upper(), dept_raw.upper() if dept_raw else ""
                is_notation_err = False
                
                if agency_upper not in MAP_CHUAN_HOA_LOAI:
                    error_list.append(f"❌ Số hiệu sai loại văn bản: `[{agency_raw}]` → Không thuộc 27 loại quy định.")
                    is_notation_err = True
                else:
                    correct_agency_case = MAP_CHUAN_HOA_LOAI[agency_upper]
                    if agency_raw != correct_agency_case:
                        error_list.append(f"❌ Sai quy tắc viết hoa loại VB: `[{agency_raw}]` → Phải viết là **{correct_agency_case}** (Ví dụ: TTr[cite: 5]).")
                        is_notation_err = True

                if dept == "":
                    is_notation_err = True
                    correct_agency = MAP_CHUAN_HOA_LOAI.get(agency_upper, agency_raw)
                    choices = [f"Chỉ {correct_agency} (Bỏ dấu gạch ngang)"] + [f"{pb['ten']} ({pb['viet_tat']})" for pb in HIDDEN_CONFIG["danh_sach_phong_ban"]]
                    ambiguous_dict[full_notation] = choices
                    error_list.append(f"⚠️ Số hiệu trống/thiếu phòng ban: `[{full_notation}]` → Cần bổ sung thông tin.")
                else:
                    if notation_match.group(2) != dept:
                        error_list.append(f"❌ Viết thường đơn vị: `[{dept_raw}]` → Phải viết HOA toàn bộ (**{dept}**).")
                        is_notation_err = True
                    if dept not in valid_vts:
                        is_notation_err = True
                        correct_agency = MAP_CHUAN_HOA_LOAI.get(agency_upper, agency_raw)
                        possible_matches = [vt for vt in valid_vts if vt.startswith(dept)]
                        if len(possible_matches) == 1:
                            error_list.append(f"❌ Số hiệu sai phòng ban: `[{full_notation}]` → Gợi ý: **/{correct_agency}-{possible_matches[0]}**")
                        else:
                            error_list.append(f"❌ Số hiệu không khớp danh mục BAWACO: `[{full_notation}]`")
                if not is_notation_err:
                    success_items.append(f"Số ký hiệu văn bản chính thức ({full_notation} - Đạt chuẩn quy định)")

    if "hành chánh" in doc_text.lower():
        error_list.append("❌ Sai từ ngữ hành chính: phát hiện từ **hành chánh**.")
    return success_items, error_list, warning_list, ambiguous_dict

# --- GIAO DIỆN CHÍNH ---
st.title("KIỂM TRA THỂ THỨC VĂN BẢN (NĐ 30)")
st.markdown("🚀 **Phiên bản V5.0:** Đã tích hợp cấu phần quét **Căn lề (Top/Bottom/Left/Right)** và cưỡng ép phông nền **Times New Roman** chuẩn hóa.")
st.markdown("---")

with st.sidebar:
    st.markdown("### 📁 BỘ NẠP VĂN BẢN")
    uploaded_file = st.file_uploader("Chọn file bản thảo văn bản (.docx):", type=["docx"])
    st.markdown("---")
    st.markdown("### 🏢 DANH MỤC BAWACO")
    with st.expander("Xem bảng viết tắt chuẩn ngầm", expanded=False):
        st.caption(f"**Công ty:** {HIDDEN_CONFIG['ten_co_quan_me']}")
        for pb in HIDDEN_CONFIG["danh_sach_phong_ban"]:
            st.caption(f"• **{pb['viet_tat']}**: {pb['ten']}")

user_resolutions = {}

if uploaded_file is not None:
    doc = Document(uploaded_file)
    success_items, error_list, warning_list, ambiguous_dict = analyze_document_v5(doc)
    
    st.markdown("### 📊 Kết quả Phân tích & Kiểm tra toàn diện")
    col_overview, col_highlight = st.columns([1, 1])
    
    with col_overview:
        st.write("**📂 Đang xử lý file:** `" + uploaded_file.name + "`")
        with st.expander("Các tiêu chí đã đạt chuẩn", expanded=True):
            for item in success_items: st.write(f"✔️ {item}", unsafe_allow_html=True)
            
    with col_highlight:
        st.markdown("### 🖍️ DANH SÁCH LỖI / SAI LỆCH THỂ THỨC")
        if error_list:
            with st.chat_message("assistant", avatar="📝"):
                for item in error_list: st.markdown(item, unsafe_allow_html=True)
        else:
            st.success("✨ Xuất sắc! Văn bản đạt điểm tối đa, không phát hiện lỗi lề hay font!")
        
        if ambiguous_dict:
            for notation, choices in ambiguous_dict.items():
                user_resolutions[notation] = st.selectbox(f"Chọn đuôi đúng cho '{notation}':", options=choices)

    # --- HỘP CÔNG CỤ XỬ LÝ (AUTOFIX NÂNG CẤP V5.0) ---
    st.markdown("### 🛠️ HỘP CÔNG CỤ XỬ LÝ TỰ ĐỘNG CHUẨN HOÁ")
    if st.button("🪄 TỰ ĐỘNG FIX TOÀN DIỆN (LỀ + FONT + CHỮ)", type="primary"):
        with st.spinner("Đang định dạng lại lề trang, phông chữ và nội dung..."):
            
            # Fix 1: Ép lề chuẩn Nghị định 30 (Lấy cận giữa tối ưu nhất)
            for section in doc.sections:
                section.top_margin = Cm(2.0)
                section.bottom_margin = Cm(2.0)
                section.left_margin = Cm(3.0)
                section.right_margin = Cm(1.5)

            # Fix 2: Duyệt toàn bộ Paragraph để ép Font và nội dung
            def fix_para(p, paragraph_index):
                text_clean = p.text.replace("|", "").strip()
                if not text_clean: return
                text_upper = unicodedata.normalize('NFC', text_clean.upper())
                orig_size = p.runs[0].font.size if p.runs and p.runs[0].font.size else Pt(13)

                # ==========================================
                # KHỐI 1: XỬ LÝ MẠNH TAY PHẦN HEADER (15 dòng đầu)
                # ==========================================
                if paragraph_index < 1:
                    
                    # 1. Cơ quan chủ quản / Đơn vị ban hành (ÉP VIẾT HOA)
                    # Lưới lọc mở rộng: Bắt mọi từ khóa tổ chức, né các từ khóa của nội dung khác
                    if any(x in text_upper for x in ["CÔNG TY", "CẤP NƯỚC", "PHÒNG", "BAN", "XÍ NGHIỆP", "TRUNG TÂM", "ỦY BAN", "UBND", "SỞ"]):
                        if not any(x in text_upper for x in ["CỘNG HÒA", "ĐỘC LẬP", "SỐ:", "NGÀY", "THÁNG", "NĂM", "CĂN CỨ", "CHỦ TỊCH", "GIÁM ĐỐC", "BAN QUẢN LÝ","KÍNH GỬI"]):
                            
                            hoa_text = text_clean.upper()
                            was_bold = any(r.bold for r in p.runs)
                            is_dept = any(x in text_upper for x in ["PHÒNG", "XÍ NGHIỆP", "TRUNG TÂM", "ĐỘI"])
                            
                            # Xóa sạch nền cũ và ép chữ in hoa
                            p.text = "" 
                            r = p.add_run(hoa_text)
                            r.font.size = Pt(13)
                            set_font_times(r)
                            
                            # Tự động in đậm cho đơn vị ban hành (Thường là cấp Phòng hoặc chính Công ty)
                            if is_dept or was_bold: 
                                r.bold = True
                            else:
                                r.bold = True 
                            p.alignment = 1 # Căn giữa
                            return 

                    # 2. Quốc hiệu & Tiêu ngữ
                    if "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM" in text_upper or re.search(r"CỘNG\s*H[ÒOÀA]+\s*XÃ\s*HỘI", text_upper):
                        p.text = ""
                        r = p.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
                        r.bold = True; r.font.size = orig_size; set_font_times(r)
                        p.alignment = 1
                        return
                    if "độc lập" in text_clean.lower() and "hạnh phúc" in text_clean.lower():
                        p.text = ""
                        r = p.add_run("Độc lập - Tự do - Hạnh phúc")
                        r.bold = True; r.font.size = Pt(14); set_font_times(r)
                        # Can thiệp XML để ép đường gạch chân mảnh
                        rPr = r._r.get_or_add_rPr()
                        u = OxmlElement('w:u')
                        u.set(qn('w:val'), 'single') 
                        u.set(qn('w:sz'), '0.1')
                        u.set(qn('w:space'), '12')  
                        # Ép kiểu nét đơn mảnh, không bị đậm theo chữ
                        rPr.append(u)
                        p.alignment = 1
                        return

                    # 3. Địa danh và Ngày tháng
                    text_lower = text_clean.lower()
                    if "ngày" in text_lower and "tháng" in text_lower and "năm" in text_lower:
                        if len(text_clean) < 70 and not any(x in text_lower for x in ["căn cứ", "luật", "nghị định", "quyết định", "thông tư", "v/v", "về việc"]):
                            p.text = "Cà Mau, ngày    tháng    năm 2026"
                            for r in p.runs: 
                                r.italic = True
                                r.font.size = Pt(13)
                                set_font_times(r)
                            p.alignment = 1
                            return

                    # 4. Số hiệu 
                    if re.match(r"^\s*Số\s*:", text_clean, re.IGNORECASE) or "Số:" in text_clean:
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
                            except NameError:
                                pass
                                
                            dt = dt_raw.upper() if dt_raw else ""
                            if dt in valid_vts_upper: return f"/{correct_agency}-{valid_vts[valid_vts_upper.index(dt)]}"
                            return f"/{correct_agency}-{dt}" if dt else f"/{correct_agency}"

                        temp_text = re.sub(r"/\s*([A-ZĐa-zđ0-9]+)(?:\s*-\s*([A-ZĐa-zđ0-9]*))?", fix_notation, temp_text)
                        if temp_text != p.text:
                            was_bold = any(r.bold for r in p.runs)
                            p.text = temp_text
                            for r in p.runs:
                                r.bold = was_bold; r.font.size = orig_size; set_font_times(r)
                        return

                # ==========================================
                # KHỐI 2: NỘI DUNG CHÍNH & FOOTER
                # ==========================================
                for r in p.runs:
                    if not r.text.strip(): continue
                    
                    if re.search(r"\bhành\s+chánh\b", r.text, flags=re.IGNORECASE):
                        r.text = re.sub(r"\bhành\s+chánh\b", "hành chính", r.text, flags=re.IGNORECASE)
                    
                    for c in ["Bạc Liêu", "Cà Mau", "Hà Nội", "Hồ Chí Minh"]:
                        if c.upper() not in r.text:
                            pattern = re.compile(rf"\b{re.escape(c)}\b", re.IGNORECASE)
                            r.text = pattern.sub(c, r.text)
                    
                    set_font_times(r)

            for idx, p in enumerate(doc.paragraphs): fix_para(p, idx)
            # Sửa tương tự cho các ô trong bảng:
            for t in doc.tables:
                for row in t.rows:
                    for cell in row.cells:
                        for idx, p in enumerate(cell.paragraphs): # Thêm enumerate và idx
                            fix_para(p, idx)

            fixed_buffer = io.BytesIO()
            doc.save(fixed_buffer)
            fixed_buffer.seek(0)
            
            st.success("🎉 HOÀN THÀNH: Đã tự động căn lề chuẩn 2-2-3-1.5 cm và nắn toàn bộ văn bản về font Times New Roman!")
            st.download_button(label="📥 TẢI VỀ FILE SỬA ĐỔI HOÀN HẢO (V5.0)", data=fixed_buffer, file_name="Fixed_V5.0_" + uploaded_file.name)
else:
    st.info("💡 Hướng dẫn: Mở thanh SLIDE bên trái và tải file bản thảo lên để bắt đầu quét lỗi nâng cao.")
