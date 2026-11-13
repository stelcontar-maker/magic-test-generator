import streamlit as st
import fitz  # PyMuPDF，用於處理與裁切 PDF 影像
import docx
from docx.shared import Inches
import io
import re

# 網頁基本設定
st.set_page_config(page_title="備課魔法師 - Word 試卷產生器", page_icon="🪄", layout="centered")

def parse_markdown(md_text):
    """解析 Markdown/TXT，精準抓取每題的標題、內容與解答"""
    # 以【第 X 頁】作為切分點
    pages = re.split(r'【第 \d+ 頁】', md_text)
    questions = []
    
    for page in pages:
        # 只抓取「標題：」到「【解答：X】」之間的所有內容，過濾掉底下的排版與提示詞
        match = re.search(r'(標題：.*?【解答：.】)', page, re.DOTALL)
        if match:
            q_text = match.group(1).strip()
            questions.append(q_text)
            
    return questions

def create_word_doc(questions, pdf_bytes):
    """將文字與裁切後的 PDF 圖片組裝成 Word"""
    doc = docx.Document()
    doc.add_heading("備課魔法師 - 完美素養測驗卷", 0)
    
    # 將上傳的 PDF 讀入記憶體
    pdf_doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    
    for i, q_text in enumerate(questions):
        doc.add_heading(f"【第 {i+1} 題】", level=2)
        
        # 1. 寫入文字區塊
        lines = q_text.split('\n')
        for line in lines:
            if line.strip():
                doc.add_paragraph(line.strip())
        
        # 2. 擷取並裁切圖片
        if i < len(pdf_doc):
            page = pdf_doc[i]
            rect = page.rect
            # 設定裁切範圍：只保留右側 45% 的區域 (避開左側文字，只抓生物圖表)
            # 參數：左上角 X, 左上角 Y, 右下角 X, 右下角 Y
            clip_rect = fitz.Rect(rect.width * 0.55, rect.height * 0.1, rect.width, rect.height * 0.9)
            
            # 以高畫質 (dpi=200) 輸出圖片
            pix = page.get_pixmap(clip=clip_rect, dpi=200)
            img_stream = io.BytesIO(pix.tobytes("png"))
            
            # 3. 寫入圖片
            doc.add_picture(img_stream, width=Inches(3.5))
        
        # 題與題之間加入分隔線
        doc.add_paragraph("-" * 50)
        
    return doc

# ==========================================
# 網頁前端介面
# ==========================================
st.title("🪄 備課魔法師 - 完美 Word 試卷產生器")
st.markdown("將生成的 **文字檔 (.txt/.md)** 與 **排版好的簡報 (.pdf)** 同時上傳，系統會自動萃取文字並裁切圖片，組裝成完美的 Word 考卷。")

# 建立兩個檔案上傳區塊
col1, col2 = st.columns(2)
with col1:
    uploaded_md = st.file_uploader("1. 上傳題庫文字檔 (TXT/MD)", type=['txt', 'md'])
with col2:
    # 這裡限定 PDF 是因為 PDF 的底層影像裁切最為精準且快速
    uploaded_pdf = st.file_uploader("2. 上傳含圖測驗卷 (PDF)", type=['pdf'])

if uploaded_md and uploaded_pdf:
    st.success("檔案讀取成功！點擊下方按鈕開始轉換。")
    
    if st.button("🚀 開始合成 Word 試卷", use_container_width=True):
        with st.spinner("正在解析文字與裁切圖片，請稍候..."):
            try:
                # 讀取文字檔內容 (處理中文編碼)
                md_text = uploaded_md.getvalue().decode("utf-8")
                questions = parse_markdown(md_text)
                
                # 讀取 PDF 二進位檔
                pdf_bytes = uploaded_pdf.getvalue()
                
                # 執行合併轉換
                final_doc = create_word_doc(questions, pdf_bytes)
                
                # 將產出的 Word 存入記憶體準備下載
                word_io = io.BytesIO()
                final_doc.save(word_io)
                word_io.seek(0)
                
                st.balloons()
                st.download_button(
                    label="📥 下載可編輯的 Word 考卷",
                    data=word_io,
                    file_name="備課魔法師_素養考卷.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"轉換過程中發生錯誤：{e}")
