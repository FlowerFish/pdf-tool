import streamlit as st
import pdfplumber
import fitz  # PyMuPDF
from PIL import Image
import os
import io
import base64
import tempfile
import zipfile

class PDFProcessor:
    @staticmethod
    def split_pdf(pdf_bytes, mode="single"):
        try:
            if not pdf_bytes:
                return False, "輸入的 PDF 數據為空"

            temp_dir = tempfile.mkdtemp()
            split_files = []
            
            # 將上傳的 PDF 寫入臨時文件
            temp_pdf_path = os.path.join(temp_dir, "temp.pdf")
            with open(temp_pdf_path, "wb") as f:
                f.write(pdf_bytes)
            
            # 驗證臨時檔案是否正確寫入
            if not os.path.exists(temp_pdf_path) or os.path.getsize(temp_pdf_path) == 0:
                return False, "臨時 PDF 檔案無效或為空"
            
            doc = fitz.open(temp_pdf_path)
            if mode == "single":
                for page_num in range(len(doc)):
                    new_pdf = fitz.open()
                    new_pdf.insert_pdf(doc, from_page=page_num, to_page=page_num)
                    output_path = os.path.join(temp_dir, f"page_{page_num+1}.pdf")
                    new_pdf.save(output_path)
                    new_pdf.close()
                    split_files.append(output_path)
            elif mode == "odd":
                for page_num in range(0, len(doc), 2):
                    new_pdf = fitz.open()
                    new_pdf.insert_pdf(doc, from_page=page_num, to_page=page_num)
                    output_path = os.path.join(temp_dir, f"page_{page_num+1}_odd.pdf")
                    new_pdf.save(output_path)
                    new_pdf.close()
                    split_files.append(output_path)
            elif mode == "even":
                for page_num in range(1, len(doc), 2):
                    new_pdf = fitz.open()
                    new_pdf.insert_pdf(doc, from_page=page_num, to_page=page_num)
                    output_path = os.path.join(temp_dir, f"page_{page_num+1}_even.pdf")
                    new_pdf.save(output_path)
                    new_pdf.close()
                    split_files.append(output_path)
            doc.close()
            
            # 創建壓縮包
            zip_path = os.path.join(temp_dir, "split_pdfs.zip")
            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED, allowZip64=True) as zipf:
                for file in split_files:
                    if os.path.exists(file) and os.path.getsize(file) > 0:
                        zipf.write(file, os.path.basename(file))
            
            # 驗證壓縮包
            if not os.path.exists(zip_path) or os.path.getsize(zip_path) == 0:
                return False, "壓縮包未創建或為空"
            
            # 讀取壓縮包
            with open(zip_path, "rb") as f:
                zip_data = f.read()
            
            return True, zip_data
        except Exception as e:
            return False, str(e)

    @staticmethod
    def extract_pages(pdf_bytes, pages_input):
        try:
            if not pdf_bytes:
                return False, "輸入的 PDF 數據為空"

            temp_dir = tempfile.mkdtemp()
            temp_pdf_path = os.path.join(temp_dir, "temp.pdf")
            with open(temp_pdf_path, "wb") as f:
                f.write(pdf_bytes)
            
            # 驗證臨時檔案
            if not os.path.exists(temp_pdf_path) or os.path.getsize(temp_pdf_path) == 0:
                return False, "臨時 PDF 檔案無效或為空"
                
            doc = fitz.open(temp_pdf_path)
            new_pdf = fitz.open()
            page_list = []
            for part in pages_input.split(","):
                if "-" in part:
                    start, end = map(int, part.split("-"))
                    page_list.extend(range(start-1, end))
                else:
                    page_list.append(int(part)-1)
            page_list = sorted(set(p for p in page_list if 0 <= p < len(doc)))
            for page_num in page_list:
                new_pdf.insert_pdf(doc, from_page=page_num, to_page=page_num)
                
            output_path = os.path.join(temp_dir, "extracted_pages.pdf")
            new_pdf.save(output_path)
            new_pdf.close()
            doc.close()
            
            with open(output_path, "rb") as f:
                output_pdf_bytes = f.read()
                
            return True, output_pdf_bytes
        except Exception as e:
            return False, str(e)

    @staticmethod
    def merge_pdfs(pdf_files):
        try:
            temp_dir = tempfile.mkdtemp()
            result_pdf = fitz.open()
            
            for pdf_file in pdf_files:
                pdf_bytes = pdf_file.read()
                if not pdf_bytes:
                    continue
                temp_path = os.path.join(temp_dir, pdf_file.name)
                with open(temp_path, "wb") as f:
                    f.write(pdf_bytes)
                doc = fitz.open(temp_path)
                result_pdf.insert_pdf(doc)
                doc.close()
                
            output_path = os.path.join(temp_dir, "merged.pdf")
            result_pdf.save(output_path)
            result_pdf.close()
            
            with open(output_path, "rb") as f:
                output_pdf_bytes = f.read()
                
            return True, output_pdf_bytes
        except Exception as e:
            return False, str(e)
            
    @staticmethod
    def extract_images(pdf_bytes, selected_indices, image_format="png"):
        try:
            if not pdf_bytes:
                return False, [], "輸入的 PDF 數據為空"

            temp_dir = tempfile.mkdtemp()
            temp_pdf_path = os.path.join(temp_dir, "temp.pdf")
            with open(temp_pdf_path, "wb") as f:
                f.write(pdf_bytes)
            
            # 驗證臨時檔案
            if not os.path.exists(temp_pdf_path) or os.path.getsize(temp_pdf_path) == 0:
                return False, [], "臨時 PDF 檔案無效或為空"
            
            doc = fitz.open(temp_pdf_path)
            all_images = []
            
            # 提取所有圖片
            for page_index in range(len(doc)):
                for img_index, img in enumerate(doc[page_index].get_images(full=True)):
                    xref = img[0]
                    base_image = doc.extract_image(xref)
                    image_bytes = base_image["image"]
                    all_images.append(image_bytes)
            
            doc.close()
            
            # 創建壓縮包
            if selected_indices:
                zip_path = os.path.join(temp_dir, "extracted_images.zip")
                with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED, allowZip64=True) as zipf:
                    for idx in selected_indices:
                        if idx < len(all_images):
                            img_bytes = all_images[idx]
                            img = Image.open(io.BytesIO(img_bytes))
                            img_path = os.path.join(temp_dir, f"image_{idx+1}.{image_format.lower()}")
                            save_format = "JPEG" if image_format.lower() == "jpg" else image_format.upper()
                            img.save(img_path, format=save_format)
                            if os.path.exists(img_path) and os.path.getsize(img_path) > 0:
                                zipf.write(img_path, os.path.basename(img_path))
                
                # 驗證壓縮包
                if not os.path.exists(zip_path) or os.path.getsize(zip_path) == 0:
                    return False, all_images, "壓縮包未創建或為空"
                
                # 讀取壓縮包
                with open(zip_path, "rb") as f:
                    zip_data = f.read()
                
                return True, all_images, zip_data
            
            return True, all_images, None
        except Exception as e:
            return False, [], str(e)

    @staticmethod
    def convert_pdf_to_txt(pdf_bytes, display_page_numbers):
        try:
            if not pdf_bytes:
                return False, "輸入的 PDF 數據為空"

            temp_dir = tempfile.mkdtemp()
            temp_pdf_path = os.path.join(temp_dir, "temp.pdf")
            with open(temp_pdf_path, "wb") as f:
                f.write(pdf_bytes)
            
            # 驗證臨時檔案
            if not os.path.exists(temp_pdf_path) or os.path.getsize(temp_pdf_path) == 0:
                return False, "臨時 PDF 檔案無效或為空"
            
            text_content = ""
            with pdfplumber.open(temp_pdf_path) as pdf:
                for page_number, page in enumerate(pdf.pages, start=1):
                    page_text = page.extract_text()
                    if page_text:
                        if display_page_numbers:
                            text_content += f"--- Page {page_number} ---\n"
                        text_content += page_text + "\n\n"
                    
                    tables = page.extract_tables()
                    for table in tables:
                        text_content += f"--- Table on Page {page_number} ---\n"
                        for row in table:
                            text_content += " | ".join(str(cell) for cell in row) + "\n"
                        text_content += "\n"
            
            return True, text_content
        except Exception as e:
            return False, str(e)

def get_binary_file_downloader_html(bin_data, file_label='File', file_name='file.zip'):
    bin_str = base64.b64encode(bin_data).decode()
    href = f'<a href="data:application/octet-stream;base64,{bin_str}" download="{file_name}">{file_label}</a>'
    return href

def main():
    st.set_page_config(
        page_title="PDF多功能工具",
        page_icon="📄",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    st.markdown("""
    <style>
    .stApp {
        max-width: 1200px;
        margin: 0 auto;
    }
    .main-header {
        font-size: 2.5rem;
        margin-bottom: 1rem;
        color: #1E3A8A;
        text-align: center;
    }
    .sub-header {
        font-size: 1.5rem;
        margin-top: 2rem;
        margin-bottom: 1rem;
        color: #2563EB;
    }
    .info-text {
        font-size: 0.9rem;
        color: #6B7280;
    }
    .success-message {
        padding: 1rem;
        background-color: #D1FAE5;
        border-radius: 0.5rem;
        color: #065F46;
    }
    .error-message {
        padding: 1rem;
        background-color: #FEE2E2;
        border-radius: 0.5rem;
        color: #B91C1C;
    }
    .image-preview {
        border: 1px solid #E5E7EB;
        border-radius: 0.5rem;
        padding: 0.5rem;
        margin: 0.5rem;
        text-align: center;
    }
    </style>
    """, unsafe_allow_html=True)
    
    st.markdown("<h1 class='main-header'>PDF多功能工具</h1>", unsafe_allow_html=True)
    
    activities = [
        "PDF轉文字", 
        "提取PDF圖片", 
        "分割PDF", 
        "擷取特定頁面", 
        "合併多個PDF"
    ]
    
    with st.sidebar:
        st.image("https://www.svgrepo.com/show/374049/pdf.svg", width=100)
        choice = st.selectbox("選擇功能", activities)
        st.markdown("---")
        st.markdown("<p class='info-text'>Stream 版本 Ver: 1.1s 2025<br>作者：葉春華</p>", unsafe_allow_html=True)
    
    if choice == "PDF轉文字":
        st.markdown("<h2 class='sub-header'>PDF轉換為文字</h2>", unsafe_allow_html=True)
        display_page_numbers = st.checkbox("顯示頁碼", value=False)
        uploaded_file = st.file_uploader("上傳PDF檔案", type=['pdf'], key="txt_uploader")
        
        if uploaded_file is not None:
            pdf_bytes = uploaded_file.read()
            with st.spinner("正在處理中..."):
                success, result = PDFProcessor.convert_pdf_to_txt(pdf_bytes, display_page_numbers)
                
                if success:
                    st.markdown("<div class='success-message'>PDF成功轉換為文字!</div>", unsafe_allow_html=True)
                    st.download_button(
                        label="下載文字檔案",
                        data=result,
                        file_name=uploaded_file.name.replace('.pdf', '.txt'),
                        mime="text/plain"
                    )
                    with st.expander("預覽文字內容"):
                        st.text_area("", result, height=300)
                else:
                    st.markdown(f"<div class='error-message'>轉換失敗: {result}</div>", unsafe_allow_html=True)
    
    elif choice == "提取PDF圖片":
        st.markdown("<h2 class='sub-header'>PDF轉換為文字</h2>", unsafe_allow_html=True)
        uploaded_file = st.file_uploader("上傳PDF檔案", type=['pdf'], key="img_uploader")
        
        if uploaded_file is not None:
            pdf_bytes = uploaded_file.read()
            image_format = st.radio("選擇圖片格式", ["PNG", "JPG"])
            
            with st.spinner("正在提取圖片..."):
                success, images, _ = PDFProcessor.extract_images(pdf_bytes, [], image_format.lower())
                
                if success and images:
                    st.markdown("<div class='success-message'>成功提取圖片!</div>", unsafe_allow_html=True)
                    
                    st.write(f"找到 {len(images)} 張圖片")
                    selected_images = []
                    
                    cols = st.columns(3)
                    for i, img_bytes in enumerate(images):
                        col_idx = i % 3
                        with cols[col_idx]:
                            st.markdown(f"<div class='image-preview'>", unsafe_allow_html=True)
                            img = Image.open(io.BytesIO(img_bytes))
                            img.thumbnail((200, 200))
                            st.image(img, caption=f"圖片 {i+1}")
                            selected = st.checkbox(f"選擇圖片 {i+1}", key=f"img_{i}")
                            if selected:
                                selected_images.append(i)
                            st.markdown("</div>", unsafe_allow_html=True)
                    
                    if st.button("下載選中的圖片"):
                        if selected_images:
                            with st.spinner("正在準備下載..."):
                                success, _, zip_data = PDFProcessor.extract_images(
                                    pdf_bytes, 
                                    selected_images, 
                                    image_format.lower()
                                )
                                if success:
                                    st.download_button(
                                        label="下載圖片壓縮包",
                                        data=zip_data,
                                        file_name="extracted_images.zip",
                                        mime="application/zip"
                                    )
                                else:
                                    st.markdown("<div class='error-message'>生成壓縮包失敗</div>", unsafe_allow_html=True)
                        else:
                            st.warning("請至少選擇一張圖片")
                elif success and not images:
                    st.warning("未在PDF中找到任何圖片")
                else:
                    st.markdown(f"<div class='error-message'>提取失敗: {images}</div>", unsafe_allow_html=True)
    
    elif choice == "分割PDF":
        st.markdown("<h2 class='sub-header'>分割PDF</h2>", unsafe_allow_html=True)
        uploaded_file = st.file_uploader("上傳PDF檔案", type=['pdf'], key="split_uploader")
        
        if uploaded_file is not None:
            pdf_bytes = uploaded_file.read()
            mode = st.radio("選擇分割模式", ["分割成單頁", "擷取單數頁", "擷取偶數頁"], 
                            format_func=lambda x: {"分割成單頁": "分割成單頁", "擷取單數頁": "擷取單數頁", "擷取偶數頁": "擷取偶數頁"}[x])
            
            mode_map = {
                "分割成單頁": "single",
                "擷取單數頁": "odd",
                "擷取偶數頁": "even"
            }
            
            if st.button("分割PDF"):
                with st.spinner("正在處理中..."):
                    success, result = PDFProcessor.split_pdf(pdf_bytes, mode_map[mode])
                    
                    if success:
                        st.markdown("<div class='success-message'>PDF成功分割!</div>", unsafe_allow_html=True)
                        st.download_button(
                            label="下載分割後的PDF壓縮包",
                            data=result,
                            file_name="split_pdfs.zip",
                            mime="application/zip"
                        )
                    else:
                        st.markdown(f"<div class='error-message'>分割失敗: {result}</div>", unsafe_allow_html=True)
    
    elif choice == "擷取特定頁面":
        st.markdown("<h2 class='sub-header'>擷取特定頁面</h2>", unsafe_allow_html=True)
        uploaded_file = st.file_uploader("上傳PDF檔案", type=['pdf'], key="extract_uploader")
        
        if uploaded_file is not None:
            pdf_bytes = uploaded_file.read()
            pages_input = st.text_input("輸入頁面範圍 (例如: 1-3,5,7-9)")
            
            if st.button("擷取頁面"):
                if pages_input:
                    with st.spinner("正在處理中..."):
                        success, result = PDFProcessor.extract_pages(pdf_bytes, pages_input)
                        
                        if success:
                            st.markdown("<div class='success-message'>成功擷取指定頁面!</div>", unsafe_allow_html=True)
                            st.download_button(
                                label="下載擷取的PDF",
                                data=result,
                                file_name="extracted_pages.pdf",
                                mime="application/pdf"
                            )
                        else:
                            st.markdown(f"<div class='error-message'>擷取失敗: {result}</div>", unsafe_allow_html=True)
                else:
                    st.warning("請輸入頁面範圍")
    
    elif choice == "合併多個PDF":
        st.markdown("<h2 class='sub-header'>合併多個PDF</h2>", unsafe_allow_html=True)
        uploaded_files = st.file_uploader("上傳多個PDF檔案", type=['pdf'], accept_multiple_files=True, key="merge_uploader")
        
        if uploaded_files:
            st.write(f"已上傳 {len(uploaded_files)} 個PDF檔案")
            file_names = [file.name for file in uploaded_files]
            for name in file_names:
                st.text(f"• {name}")
            
            if st.button("合併PDF"):
                with st.spinner("正在合併中..."):
                    success, result = PDFProcessor.merge_pdfs(uploaded_files)
                    
                    if success:
                        st.markdown("<div class='success-message'>PDF成功合併!</div>", unsafe_allow_html=True)
                        st.download_button(
                            label="下載合併後的PDF",
                            data=result,
                            file_name="merged.pdf",
                            mime="application/pdf"
                        )
                    else:
                        st.markdown(f"<div class='error-message'>合併失敗: {result}</div>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
