import streamlit as st
import fitz  # PyMuPDF
from pptx import Presentation
from pptx.util import Inches, Pt
from io import BytesIO

def check_password():
    # 這裡放入你之前設定的密碼邏輯
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False
    if not st.session_state["password_correct"]:
        pwd = st.text_input("請輸入管理員密碼", type="password")
        if st.button("確認登入"):
            if pwd == "54167":
                st.session_state["password_correct"] = True
                st.rerun()
        return False
    return True

def convert_pdf_to_simple_pptx(uploaded_file):
    doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
    prs = Presentation()
    # 設定 16:9
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    for page in doc:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # --- 1. 背景圖片層 ---
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        img_data = pix.tobytes("png")
        slide.shapes.add_picture(BytesIO(img_data), 0, 0, width=prs.slide_width, height=prs.slide_height)

        # --- 2. 文字層優化：按「行」合併 ---
        # 使用 "dict" 模式獲取結構化數據
        page_dict = page.get_text("dict")
        page_w = page.rect.width
        page_h = page.rect.height

        for block in page_dict["blocks"]:
            if "lines" in block:
                for line in block["lines"]:
                    # 合併這一行所有的 spans (文字片段)
                    full_line_text = "".join([span["text"] for span in line["spans"]])
                    if not full_line_text.strip(): continue
                    
                    # 取得這一行的邊界
                    bbox = line["bbox"] # (x0, y0, x1, y1)
                    
                    # 轉換為 PPT 座標
                    x = (bbox[0] / page_w) * prs.slide_width
                    y = (bbox[1] / page_h) * prs.slide_height
                    w = ((bbox[2] - bbox[0]) / page_w) * prs.slide_width
                    h = ((bbox[3] - bbox[1]) / page_h) * prs.slide_height
                    
                    # 取得該行第一個片段的字體大小作為基準
                    base_font_size = line["spans"][0]["size"]

                    # 在圖片上方建立文字框
                    txBox = slide.shapes.add_textbox(x, y, w, h)
                    tf = txBox.text_frame
                    tf.text = full_line_text
                    
                    # 設定字體樣式
                    p = tf.paragraphs[0]
                    p.font.size = Pt(base_font_size * 0.8) # 縮放系數微調
                    # 讓文字框背景透明（PPT 預設通常是透明的）

    ppt_output = BytesIO()
    prs.save(ppt_output)
    return ppt_output.getvalue()

# --- 主介面 ---
if check_password():
    st.title("🚀 精簡版圖文分離工具")
    st.write("目標：背景圖一層 + 每行文字各一個框，不再碎碎的。")
    
    file = st.file_uploader("上傳 PDF", type="pdf")
    if file and st.button("開始轉換"):
        with st.spinner("正在提取圖層..."):
            result = convert_pdf_to_simple_pptx(file)
            st.download_button("📥 下載 PPTX", result, file_name="Simple_Layout.pptx")
