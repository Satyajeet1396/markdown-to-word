import streamlit as st
import requests
from io import BytesIO
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import re
import base64

# Page configuration
st.set_page_config(
    page_title="Markdown to Word Converter",
    page_icon="📄",
    layout="wide"
)

# Title and description
st.title("📄 Markdown to Word Converter")
st.markdown("Convert markdown content from AI websites or GitHub into formatted Word documents")

# Sidebar for GitHub integration
with st.sidebar:
    st.header("🔧 Settings")
    
    # GitHub section
    st.subheader("GitHub Repository")
    github_url = st.text_input(
        "GitHub File URL",
        placeholder="https://github.com/user/repo/blob/main/file.md",
        help="Enter the GitHub URL of a markdown file"
    )
    
    if st.button("📥 Fetch from GitHub"):
        if github_url:
            try:
                # Convert GitHub URL to raw URL
                raw_url = github_url.replace("github.com", "raw.githubusercontent.com").replace("/blob/", "/")
                response = requests.get(raw_url)
                response.raise_for_status()
                st.session_state['markdown_content'] = response.text
                st.success("✅ Successfully fetched from GitHub!")
            except Exception as e:
                st.error(f"❌ Error fetching from GitHub: {str(e)}")
        else:
            st.warning("Please enter a GitHub URL")
    
    st.divider()
    
    # Document settings
    st.subheader("Document Settings")
    doc_title = st.text_input("Document Title", value="Converted Document")
    font_size = st.slider("Base Font Size", 8, 16, 11)
    
    # Style options
    st.subheader("Styling Options")
    use_colors = st.checkbox("Use colored headings", value=True)
    add_toc = st.checkbox("Add table of contents", value=False)

# Main content area with tabs
tab1, tab2, tab3 = st.tabs(["📝 Input Markdown", "👁️ Preview", "⬇️ Download"])

with tab1:
    st.subheader("Paste or Edit Your Markdown Content")
    
    # Initialize session state for markdown content
    if 'markdown_content' not in st.session_state:
        st.session_state['markdown_content'] = """# Sample Markdown

## Introduction
This is a **sample** markdown document with *formatting*.

### Features
- Bullet point 1
- Bullet point 2
- Bullet point 3

### Code Example
```python
def hello_world():
    print("Hello, World!")
```

### Numbered List
1. First item
2. Second item
3. Third item

**Bold text** and *italic text* are supported.
"""
    
    markdown_input = st.text_area(
        "Markdown Content",
        value=st.session_state['markdown_content'],
        height=400,
        help="Paste markdown content from ChatGPT, Claude, or any AI website"
    )
    
    if st.button("🔄 Update Content"):
        st.session_state['markdown_content'] = markdown_input
        st.success("Content updated!")

with tab2:
    st.subheader("Markdown Preview")
    st.markdown(st.session_state['markdown_content'])

with tab3:
    st.subheader("Generate Word Document")
    
    def parse_markdown_to_docx(markdown_text, title, font_size, use_colors):
        """Convert markdown to Word document with formatting"""
        doc = Document()
        
        # Add title
        title_para = doc.add_heading(title, 0)
        title_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        
        lines = markdown_text.split('\n')
        in_code_block = False
        code_lines = []
        in_list = False
        
        for line in lines:
            # Handle code blocks
            if line.strip().startswith('```'):
                if in_code_block:
                    # End code block
                    code_text = '\n'.join(code_lines)
                    para = doc.add_paragraph(code_text)
                    para.style = 'Intense Quote'
                    for run in para.runs:
                        run.font.name = 'Courier New'
                        run.font.size = Pt(font_size - 1)
                    code_lines = []
                    in_code_block = False
                else:
                    # Start code block
                    in_code_block = True
                continue
            
            if in_code_block:
                code_lines.append(line)
                continue
            
            # Handle headings
            if line.startswith('# '):
                para = doc.add_heading(line[2:], 1)
                if use_colors:
                    for run in para.runs:
                        run.font.color.rgb = RGBColor(0, 51, 102)
            elif line.startswith('## '):
                para = doc.add_heading(line[3:], 2)
                if use_colors:
                    for run in para.runs:
                        run.font.color.rgb = RGBColor(51, 102, 153)
            elif line.startswith('### '):
                para = doc.add_heading(line[4:], 3)
                if use_colors:
                    for run in para.runs:
                        run.font.color.rgb = RGBColor(102, 153, 204)
            
            # Handle bullet lists
            elif line.strip().startswith('- ') or line.strip().startswith('* '):
                text = line.strip()[2:]
                para = doc.add_paragraph(text, style='List Bullet')
                apply_inline_formatting(para, font_size)
                in_list = True
            
            # Handle numbered lists
            elif re.match(r'^\d+\.\s', line.strip()):
                text = re.sub(r'^\d+\.\s', '', line.strip())
                para = doc.add_paragraph(text, style='List Number')
                apply_inline_formatting(para, font_size)
                in_list = True
            
            # Handle regular paragraphs
            elif line.strip():
                if in_list:
                    doc.add_paragraph()  # Add space after list
                    in_list = False
                para = doc.add_paragraph(line)
                apply_inline_formatting(para, font_size)
            
            # Handle empty lines
            else:
                if not in_list:
                    doc.add_paragraph()
        
        return doc
    
    def apply_inline_formatting(paragraph, font_size):
        """Apply bold and italic formatting to paragraph text"""
        text = paragraph.text
        paragraph.clear()
        
        # Pattern to match **bold**, *italic*, and `code`
        pattern = r'(\*\*.*?\*\*|\*.*?\*|`.*?`)'
        parts = re.split(pattern, text)
        
        for part in parts:
            if part.startswith('**') and part.endswith('**'):
                run = paragraph.add_run(part[2:-2])
                run.bold = True
                run.font.size = Pt(font_size)
            elif part.startswith('*') and part.endswith('*') and not part.startswith('**'):
                run = paragraph.add_run(part[1:-1])
                run.italic = True
                run.font.size = Pt(font_size)
            elif part.startswith('`') and part.endswith('`'):
                run = paragraph.add_run(part[1:-1])
                run.font.name = 'Courier New'
                run.font.size = Pt(font_size - 1)
                run.font.color.rgb = RGBColor(220, 50, 50)
            else:
                run = paragraph.add_run(part)
                run.font.size = Pt(font_size)
    
    if st.button("📄 Generate Word Document"):
        with st.spinner("Generating document..."):
            try:
                # Generate the document
                doc = parse_markdown_to_docx(
                    st.session_state['markdown_content'],
                    doc_title,
                    font_size,
                    use_colors
                )
                
                # Save to BytesIO
                bio = BytesIO()
                doc.save(bio)
                bio.seek(0)
                
                st.success("✅ Document generated successfully!")
                
                # Download button
                st.download_button(
                    label="⬇️ Download Word Document",
                    data=bio,
                    file_name=f"{doc_title.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
            except Exception as e:
                st.error(f"❌ Error generating document: {str(e)}")

# Footer
st.divider()
st.markdown("""
### 💡 Tips:
- Paste markdown directly from ChatGPT, Claude, Gemini, or any AI website
- Use GitHub URLs to fetch markdown files from repositories
- Supports headings, lists, bold, italic, and code blocks
- Customize font size and styling options in the sidebar
""")

# Instructions section
with st.expander("📖 How to Use"):
    st.markdown("""
    **Method 1: Direct Paste**
    1. Copy markdown content from any AI website (ChatGPT, Claude, etc.)
    2. Paste it into the "Input Markdown" tab
    3. Click "Update Content"
    4. Go to "Download" tab and click "Generate Word Document"
    
    **Method 2: GitHub Import**
    1. Copy the GitHub URL of a markdown file
    2. Paste it in the sidebar under "GitHub File URL"
    3. Click "Fetch from GitHub"
    4. Go to "Download" tab and generate the document
    
    **GitHub URL Format:**
    - `https://github.com/username/repository/blob/main/file.md`
    - The app will automatically convert it to the raw URL
    """)
