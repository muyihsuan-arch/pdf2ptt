import streamlit as st
import fitz  # PyMuPDF
from pptx import Presentation
from pptx.util import Inches, Pt
from io import BytesIO
from PIL import Image
import os

# --- 密碼保護 ---
def check_password():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if not st.session_state.authenticated:
        pwd = st.sidebar.text_input("請輸入密碼以開啟 PRO 功能", type="password")
        if pwd == "54167": # 請修改這裡
            st.session_state.authenticated = True
            st.rerun()
        else:
            if pwd: st.sidebar.error("密碼錯誤")
            st.title("🔒 存取受限")
            st.info("請輸入正確密碼以解鎖 PDF 轉 PPT PRO 工具。")
            return False
    return True

if check_password():
    st.title("🚀 PDF 轉 PPT 圖文分離版")
    st.caption("自動提取 PDF 背景圖並將文字轉為可編輯圖層 (Powered by Gemini AI 邏輯)")

    uploaded_file = st.file_uploader("上傳 PDF 檔案", type="pdf")

    if uploaded_file:
        # 讀取 PDF
        doc = fitz.open(stream=uploaded_file.read(), filetype="pdf")
        st.success(f"成功讀取: {uploaded_file.name} (共 {len(doc)} 頁)")

        if st.button("開始轉換並匯出 PPT"):
            prs = Presentation()
            # 設定 16:9 寬螢幕
            prs.slide_width = Inches(13.333)
            prs.slide_height = Inches(7.5)

            progress_bar = st.progress(0)
            
            for i, page in enumerate(doc):
                # 1. 提取背景圖 (將整頁轉為圖片)
                pix = page.get_displaylist().get_pixmap(matrix=fitz.Matrix(2, 2))
                img_data = pix.tobytes("png")
                
                # 2. 建立 PPT 投影片
                slide_layout = prs.slide_layouts[6] # 使用空白版型
                slide = prs.slides.add_slide(slide_layout)
                
                # 3. 插入背景圖 (鋪滿全螢幕)
                img_stream = BytesIO(img_data)
                slide.shapes.add_picture(img_stream, 0, 0, width=prs.slide_width, height=prs.slide_height)

                # 4. 提取文字層 (圖文分離的核心)
                # 我們把文字疊在背景圖上方，設為半透明或與原位置重合，讓使用者可以點擊編輯
                text_instances = page.get_text("dict")
                for block in text_instances["blocks"]:
                    if "lines" in block:
                        for line in block["lines"]:
                            for span in line["spans"]:
                                # 計算座標比例 (PDF 到 PPT)
                                x = (span["bbox"][0] / page.rect.width) * prs.slide_width
                                y = (span["bbox"][1] / page.rect.height) * prs.slide_height
                                w = (span["bbox"][2] - span["bbox"][0]) / page.rect.width * prs.slide_width
                                h = (span["bbox"][3] - span["bbox"][1]) / page.rect.height * prs.slide_height
                                
                                # 加入文字框
                                txBox = slide.shapes.add_textbox(x, y, w, h)
                                tf = txBox.text_frame
                                tf.text = span["text"]
                                # 嘗試匹配字體大小
                                tf.paragraphs[0].font.size = Pt(span["size"] * 0.8) 

                progress_bar.progress((i + 1) / len(doc))

            # 儲存結果
            ppt_output = BytesIO()
            prs.save(ppt_output)
            
            st.download_button(
                label="📁 下載已分離圖層的 PPTX",
                data=ppt_output.getvalue(),
                file_name=f"Converted_{uploaded_file.name}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
