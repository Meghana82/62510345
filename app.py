import streamlit as st
from markitdown import MarkItDown
import os
import io

# Initialize MarkItDown engine
# Note: MarkItDown handles Word, Excel, PPT, PDF, and HTML natively.
md = MarkItDown()

def process_file(uploaded_file):
    """Processes the uploaded file and returns Markdown text."""
    try:
        # We read the file into bytes
        file_bytes = uploaded_file.read()
        
        # MarkItDown's convert method can take file-like objects or paths.
        # Here we pass the bytes and provide the filename for extension detection.
        result = md.convert_stream(
            io.BytesIO(file_bytes), 
            file_extension=os.path.splitext(uploaded_file.name)[1]
        )
        return result.text_content
    except Exception as e:
        # Graceful error handling as per requirements
        st.warning(f"⚠️ Could not read {uploaded_file.name}. Please check the format.")
        return None

# --- UI Setup ---
st.set_page_config(page_title="Universal Doc Reader", page_icon="📄")
st.title("📄 Universal Document-to-Text Converter")
st.markdown("Upload Office docs, PDFs, or HTML files to instantly convert them to Markdown.")

# [2] Upload Area (Supports multiple files)
uploaded_files = st.file_uploader(
    "Drag and drop files here", 
    type=["docx", "xlsx", "pptx", "pdf", "html", "htm", "zip"], 
    accept_multiple_files=True
)

if uploaded_files:
    all_converted_text = ""
    
    for uploaded_file in uploaded_files:
        with st.spinner(f"Processing {uploaded_file.name}..."):
            content = process_file(uploaded_file)
            
            if content:
                # [2] Instant Preview
                st.subheader(f"Preview: {uploaded_file.name}")
                st.text_area(
                    label="Converted Text",
                    value=content,
                    height=300,
                    key=f"preview_{uploaded_file.name}"
                )
                
                # [4] Dynamic Filename Generation
                base_name = os.path.splitext(uploaded_file.name)[0]
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.download_button(
                        label=f"Download {base_name}.md",
                        data=content,
                        file_name=f"{base_name}_converted.md",
                        mime="text/markdown"
                    )
                
                with col2:
                    st.download_button(
                        label=f"Download {base_name}.txt",
                        data=content,
                        file_name=f"{base_name}_converted.txt",
                        mime="text/plain"
                    )
                
                st.divider()

else:
    st.info("Please upload a file to begin.")

# Sidebar info
with st.sidebar:
    st.header("About")
    st.write("Built with Microsoft's **MarkItDown** engine.")
    st.write("Supported formats: .docx, .xlsx, .pptx, .pdf, .html, .zip")
