import streamlit as st
import fitz  # PyMuPDF
import docx
from docx.shared import Inches
import io
import re

st.set_page_config(page_title="備課魔法師 - Word 試卷產生器", page_icon="🪄", layout="centered")

def parse_markdown(md_text):
    """解析 Markdown/TXT，精準抓取每題的標題、內容與解答"""
    pages = re.split(r'【第 \d+ 頁】', md_text)
    questions = []
    for page in pages:
        match = re.search(r'(標題：.*?【解答：.】)', page, re.DOTALL)
        if match:
            questions.append(match.group(1).strip())
    return questions

def create_word_doc(questions, pdf_bytes, crop_margins):
    """將文字與裁切後的 PDF 圖片組裝成 Word"""
    doc = docx.Document()
    doc.add_heading("備課魔法師 - 完美素養測驗卷", 0)
    
    pdf_doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    
    for i, q_text in enumerate(questions):
        doc.add_heading(f"【第 {i+1} 題】", level=2)
        
        # 1. 寫入文字
        lines = q_text.split('\n')
        for line in lines:
            if line.strip():
                doc.add_paragraph(line.strip())
        
        # 2. 擷取並依照使用者設定裁切圖片
        if i < len(pdf_doc):
            page = pdf_doc[i]
            rect = page.rect
            
            # 使用網頁滑桿傳入的比例進行裁切
            left = rect.width * (crop_margins['left'] / 100.0)
            top = rect.height * (crop_margins['top'] / 100.0)
            right = rect.width * (crop_margins['right'] / 100.0)
            bottom = rect.height * (crop_margins['bottom'] / 100.0)
            
            clip_rect = fitz.Rect(left, top, right, bottom)
            
            # 匯出圖片
            pix = page.get_pixmap(clip=clip_rect, dpi=200)
            img_stream = io.BytesIO(pix.tobytes("png"))
            
            # 3. 寫入圖片
            doc.add_picture(img_stream, width=Inches(3.5))
        
        doc.add_paragraph("-" * 50)
        
    return doc

# ==========================================
# 網頁前端介面
# ==========================================
st.title("🪄 備課魔法師 - 完美 Word 產生器")
st.markdown("將生成的 **文字檔 (.txt/.md)** 與 **測驗卷 (.pdf)** 上傳，系統會自動萃取文字並裁切圖片。")

col1, col2 = st.columns(2)
with col1:
    uploaded_md = st.file_uploader("1. 上傳題庫文字檔 (TXT/MD)", type=['txt', 'md'])
with col2:
    uploaded_pdf = st.file_uploader("2. 上傳含圖測驗卷 (PDF)", type=['pdf'])

st.divider()

# 新增：動態裁切控制區塊
st.subheader("✂️ 圖片裁切微調 (若圖片被切斷或含有文字可在此調整)")
with st.expander("展開調整裁切邊界", expanded=True):
    col_l, col_r = st.columns(2)
    with col_l:
        crop_left = st.slider("⬅️ 左側起點 (%)：數字越大切掉越多左邊文字", 0, 100, 48)
        crop_top = st.slider("⬆️ 上方起點 (%)：數字越大切掉越多上方標題", 0, 100, 20)
    with col_r:
        crop_right = st.slider("➡️ 右側終點 (%)：通常保持 100 不變", 0, 100, 100)
        crop_bottom = st.slider("⬇️ 下方終點 (%)：通常保持 95 避開頁碼", 0, 100, 95)

if uploaded_md and uploaded_pdf:
    if st.button("🚀 開始合成 Word 試卷", use_container_width=True):
        with st.spinner("正在解析文字與依照您的設定裁切圖片..."):
            try:
                md_text = uploaded_md.getvalue().decode("utf-8")
                questions = parse_markdown(md_text)
                
                crop_margins = {
                    'left': crop_left,
                    'top': crop_top,
                    'right': crop_right,
                    'bottom': crop_bottom
                }
                
                final_doc = create_word_doc(questions, uploaded_pdf.getvalue(), crop_margins)
                
                word_io = io.BytesIO()
                final_doc.save(word_io)
                word_io.seek(0)
                
                st.success("✅ 轉換成功！")
                st.download_button(
                    label="📥 下載可編輯的 Word 考卷",
                    data=word_io,
                    file_name="備課魔法師_素養考卷.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"轉換過程中發生錯誤：{e}")
