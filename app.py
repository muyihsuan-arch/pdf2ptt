import streamlit as st
import fitz  # PyMuPDF
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from io import BytesIO

# --- 1. 密碼驗證邏輯 ---
def check_password():
    """回傳 True 代表驗證通過"""
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if not st.session_state["password_correct"]:
        st.title("🔒 歡迎使用 PDF 轉 PPT PRO 工具")
        st.write("本工具僅供授權用戶使用，請先輸入密碼。")
        
        # 設定你的密碼
        password = st.text_input("請輸入密碼", type="password")
        if st.button("確認登入"):
            if password == "你的密碼": # <--- 請在這裡修改你的密碼
                st.session_state["password_correct"] = True
                st.rerun()
            else:
                st.error("密碼錯誤，請重新輸入。")
        return False
    return True

# --- 2. 核心分離邏輯 ---
def process_pdf_pro(uploaded_file):
    doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
    prs = Presentation()
    # 設置 PPT 為 16:9
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    for page in doc:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # 渲染背景圖 (確保足夠清晰)
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        img_data = pix.tobytes("png")
        slide.shapes.add_picture(BytesIO(img_data), 0, 0, width=prs.slide_width, height=prs.slide_height)

        # 提取文字並在 PPT 疊加可編輯框
        blocks = page.get_text("blocks")
        for b in blocks:
            if b[6] == 0:  # block_type 為文字
                text_content = b[4].strip()
                if not text_content: continue
                
                # 計算座標比例
                x = (b[0] / page.rect.width) * prs.slide_width
                y = (b[1] / page.rect.height) * prs.slide_height
                w = ((b[2] - b[0]) / page.rect.width) * prs.slide_width
                h = ((b[3] - b[1]) / page.rect.height) * prs.slide_height

                txBox = slide.shapes.add_textbox(x, y, w, h)
                tf = txBox.text_frame
                tf.word_wrap = True
                p = tf.paragraphs[0]
                p.text = text_content
                # 初始字體稍微設小一點，避免溢出，使用者可後續在 PPT 調整
                p.font.size = Pt(16) 
                p.alignment = PP_ALIGN.LEFT

    ppt_out = BytesIO()
    prs.save(ppt_out)
    return ppt_out.getvalue()

# --- 3. 主程式介面 ---
if check_password():
    st.set_page_config(page_title="PDF PRO Converter", layout="wide")
    
    # 側邊欄增加登出功能
    if st.sidebar.button("安全性登出"):
        st.session_state["password_correct"] = False
        st.rerun()

    st.title("🎨 DeckEdit 模倣版：圖文分離 PRO")
    st.markdown("---")

    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("第一步：上傳檔案")
        file = st.file_uploader("選擇 NotebookLM 產出的 PDF", type="pdf")
        
    with col2:
        st.subheader("第二步：執行轉換")
        if file:
            if st.button("🚀 開始深度分離圖層"):
                with st.spinner("正在解析 PDF 文字座標並提取背景..."):
                    try:
                        result = process_pdf_pro(file)
                        st.success("轉換完成！")
                        st.download_button(
                            label="📥 下載可編輯 PPTX",
                            data=result,
                            file_name=f"PRO_{file.name}.pptx",
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            use_container_width=True
                        )
                    except Exception as e:
                        st.error(f"處理失敗：{e}")
        else:
            st.info("請先上傳 PDF 檔案")
