import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from io import BytesIO

# --- 1. 密碼驗證函數 ---
def check_password():
    """回傳 True 代表密碼正確，否則回傳 False"""
    def password_entered():
        if st.session_state["password"] == "54167": # <-- 在這裡設定你的密碼
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # 清除密碼輸入框的值
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        # 第一次進入，顯示輸入框
        st.text_input("請輸入密碼以使用本工具", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        # 密碼錯誤，顯示錯誤訊息並重新輸入
        st.text_input("密碼錯誤，請再試一次", type="password", on_change=password_entered, key="password")
        st.error("😕 存取拒絕")
        return False
    else:
        # 密碼正確
        return True

# --- 2. 只有驗證通過才執行主程式 ---
if check_password():
    # --- 主程式開始 (DeckEdit 模倣版) ---
    st.set_page_config(layout="wide")
    st.title("🎨 DeckEdit 專業版 (受保護)")
    st.success("驗證成功！歡迎使用。")
    
    # 這裡放你原本的投影片生成邏輯...
    # (為了簡潔，以下省略重複的 UI 代碼，你可以直接把之前的代碼貼進來)
    
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("編輯區")
        input_text = st.text_area("請輸入 Markdown...", height=400)
    
    with col2:
        st.subheader("預覽區")
        # 預覽邏輯...
    
    # --- 登出按鈕 (選配) ---
    if st.sidebar.button("登出"):
        st.session_state["password_correct"] = False
        st.rerun()
