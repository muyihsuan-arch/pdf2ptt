import streamlit as st
import fitz  # PyMuPDF
from pptx import Presentation
from pptx.util import Inches, Pt
from io import BytesIO

# --- 密碼保護 (請記得修改你的密碼) ---
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False
    if not st.session_state["password_correct"]:
        st.title("🔐 系統鎖定")
        pwd = st.text_input("請輸入管理員密碼", type="password")
        if st.button("登入"):
            if pwd == "54167":
                st.session_state["password_correct"] = True
                st.rerun()
            else:
                st.error("密碼錯誤")
        return False
    return True

# --- 核心邏輯：中文字優化版 ---
def convert_pdf_pro_v2(uploaded_file):
    doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
    prs = Presentation()
    # 標準 16:9
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    for page in doc:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # 1. 圖片層：渲染背景 (這部分沒問題，我們維持輸出)
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        img_data = pix.tobytes("png")
        slide.shapes.add_picture(BytesIO(img_data), 0, 0, width=prs.slide_width, height=prs.slide_height)

        # 2. 文字層：解決中文抓取問題
        # 使用 "rawdict" 或 "dict" 模式，並強制抓取文本
        page_dict = page.get_text("dict")
        page_w, page_h = page.rect.width, page.rect.height

        for block in page_dict["blocks"]:
            if "lines" in block:
                for line in block["lines"]:
                    # 關鍵：將同一行內的所有中文片段(spans)強制合併
                    line_text = "".join([span["text"] for span in line["spans"]]).strip()
                    
                    if not line_text:
                        continue

                    # 取得這行文字的座標
                    bbox = line["bbox"]
                    x = (bbox[0] / page_w) * prs.slide_width
                    y = (bbox[1] / page_h) * prs.slide_height
                    w = ((bbox[2] - bbox[0]) / page_w) * prs.slide_width
                    h = ((bbox[3] - bbox[1]) / page_h) * prs.slide_height

                    # 建立文字框
                    txBox = slide.shapes.add_textbox(x, y, w, h)
                    tf = txBox.text_frame
                    tf.word_wrap = True
                    p = tf.paragraphs[0]
                    p.text = line_text
                    
                    # 嘗試抓取原始字體大小，若失敗則給預設值
                    try:
                        p.font.size = Pt(line["spans"][0]["size"] * 0.9)
                    except:
                        p.font.size = Pt(18)
                    
                    # 為了讓分離更有感，我們暫時把文字顏色設為亮色或顯色
                    # p.font.color.rgb = RGBColor(0, 0, 0)

    ppt_output = BytesIO()
    prs.save(ppt_output)
    return ppt_output.getvalue()

# --- UI 介面 ---
if check_password():
    st.title("🛠️ 中文 PDF 圖文分離 (v2 修復版)")
    st.info("如果下載後的 PPT 點擊文字可以編輯，就代表分離成功了！")
    
    file = st.file_uploader("上傳含有中文的 PDF", type="pdf")
    if file and st.button("🚀 開始執行深度分離"):
        with st.spinner("正在解析中文編碼並提取圖層..."):
            result = convert_pdf_pro_v2(file)
            st.success("分離完成！")
            st.download_button("📥 下載可編輯 PPTX", result, file_name="Separated_Chinese.pptx")
