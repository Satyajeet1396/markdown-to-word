import streamlit as st
import requests
from io import BytesIO
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
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
    preserve_math = st.checkbox("Preserve LaTeX math", value=True)

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

### Math Example
This is inline math: \\( E = mc^2 \\)

Display math:
\\[
\\int_0^\\infty e^{-x^2} dx = \\frac{\\sqrt{\\pi}}{2}
\\]
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
    
    def add_table_border(table):
        """Add borders to table"""
        tbl = table._element
        tblPr = tbl.tblPr
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr')
            tbl.insert(0, tblPr)
        
        tblBorders = OxmlElement('w:tblBorders')
        for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            border = OxmlElement(f'w:{border_name}')
            border.set(qn('w:val'), 'single')
            border.set(qn('w:sz'), '4')
            border.set(qn('w:space'), '0')
            border.set(qn('w:color'), '000000')
            tblBorders.append(border)
        tblPr.append(tblBorders)
    
    def parse_table(lines, start_idx):
        """Parse markdown table and return table data and end index"""
        table_lines = []
        idx = start_idx
        
        while idx < len(lines) and '|' in lines[idx]:
            table_lines.append(lines[idx])
            idx += 1
        
        if len(table_lines) < 2:
            return None, start_idx
        
        # Parse header
        headers = [cell.strip() for cell in table_lines[0].split('|') if cell.strip()]
        
        # Skip separator line
        rows = []
        for line in table_lines[2:]:
            cells = [cell.strip() for cell in line.split('|') if cell.strip()]
            if cells:
                rows.append(cells)
        
        return {'headers': headers, 'rows': rows}, idx
    
    def parse_markdown_to_docx(markdown_text, title, font_size, use_colors, preserve_math):
        """Convert markdown to Word document with formatting"""
        doc = Document()
        
        # Add title
        title_para = doc.add_heading(title, 0)
        title_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        
        lines = markdown_text.split('\n')
        in_code_block = False
        code_lines = []
        in_list = False
        i = 0
        
        while i < len(lines):
            line = lines[i]
            
            # Check for tables
            if '|' in line and i + 1 < len(lines) and '|' in lines[i + 1]:
                table_data, end_idx = parse_table(lines, i)
                if table_data:
                    # Add table to document
                    doc.add_paragraph()  # Space before table
                    table = doc.add_table(rows=1 + len(table_data['rows']), cols=len(table_data['headers']))
                    table.style = 'Light Grid Accent 1'
                    add_table_border(table)
                    
                    # Add headers
                    hdr_cells = table.rows[0].cells
                    for idx, header in enumerate(table_data['headers']):
                        hdr_cells[idx].text = header
                        for paragraph in hdr_cells[idx].paragraphs:
                            for run in paragraph.runs:
                                run.font.bold = True
                                run.font.size = Pt(font_size)
                    
                    # Add rows
                    for row_idx, row_data in enumerate(table_data['rows']):
                        row_cells = table.rows[row_idx + 1].cells
                        for col_idx, cell_data in enumerate(row_data):
                            if col_idx < len(row_cells):
                                row_cells[col_idx].text = cell_data
                                for paragraph in row_cells[col_idx].paragraphs:
                                    for run in paragraph.runs:
                                        run.font.size = Pt(font_size - 1)
                    
                    doc.add_paragraph()  # Space after table
                    i = end_idx
                    continue
            
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
                i += 1
                continue
            
            if in_code_block:
                code_lines.append(line)
                i += 1
                continue
            
            # Handle horizontal rules
            if line.strip() == '---':
                doc.add_paragraph('_' * 50)
                i += 1
                continue
            
            # Handle headings
            if line.startswith('# ') and not line.startswith('## '):
                para = doc.add_heading(line[2:], 1)
                if use_colors:
                    for run in para.runs:
                        run.font.color.rgb = RGBColor(0, 51, 102)
            elif line.startswith('## ') and not line.startswith('### '):
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
                apply_inline_formatting(para, font_size, preserve_math)
                in_list = True
            
            # Handle numbered lists
            elif re.match(r'^\d+\.\s', line.strip()):
                text = re.sub(r'^\d+\.\s', '', line.strip())
                para = doc.add_paragraph(text, style='List Number')
                apply_inline_formatting(para, font_size, preserve_math)
                in_list = True
            
            # Handle regular paragraphs
            elif line.strip():
                if in_list:
                    doc.add_paragraph()  # Add space after list
                    in_list = False
                para = doc.add_paragraph(line)
                apply_inline_formatting(para, font_size, preserve_math)
            
            # Handle empty lines
            else:
                if not in_list:
                    doc.add_paragraph()
            
            i += 1
        
        return doc
    
    def apply_inline_formatting(paragraph, font_size, preserve_math):
        """Apply bold, italic, code, and math formatting to paragraph text"""
        text = paragraph.text
        paragraph.clear()
        
        # Pattern to match **bold**, *italic*, `code`, \(...\) inline math, and \[...\] display math
        if preserve_math:
            pattern = r'(\\\[[\s\S]*?\\\]|\\\(.*?\\\)|\*\*.*?\*\*|\*(?!\*).*?\*(?!\*)|`.*?`)'
        else:
            pattern = r'(\*\*.*?\*\*|\*(?!\*).*?\*(?!\*)|`.*?`)'
        
        parts = re.split(pattern, text)
        
        for part in parts:
            if not part:
                continue
                
            # Display math \[...\]
            if preserve_math and part.startswith('\\[') and part.endswith('\\]'):
                run = paragraph.add_run('\n' + part + '\n')
                run.font.name = 'Cambria Math'
                run.font.size = Pt(font_size)
                run.font.color.rgb = RGBColor(0, 100, 0)
            # Inline math \(...\)
            elif preserve_math and part.startswith('\\(') and part.endswith('\\)'):
                run = paragraph.add_run(part)
                run.font.name = 'Cambria Math'
                run.font.size = Pt(font_size)
                run.font.color.rgb = RGBColor(0, 100, 0)
            # Bold **text**
            elif part.startswith('**') and part.endswith('**') and len(part) > 4:
                run = paragraph.add_run(part[2:-2])
                run.bold = True
                run.font.size = Pt(font_size)
            # Italic *text*
            elif part.startswith('*') and part.endswith('*') and len(part) > 2 and not part.startswith('**'):
                run = paragraph.add_run(part[1:-1])
                run.italic = True
                run.font.size = Pt(font_size)
            # Code `text`
            elif part.startswith('`') and part.endswith('`'):
                run = paragraph.add_run(part[1:-1])
                run.font.name = 'Courier New'
                run.font.size = Pt(font_size - 1)
                run.font.color.rgb = RGBColor(220, 50, 50)
            # Regular text
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
                    use_colors,
                    preserve_math
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
                st.error(f"Details: {type(e).__name__}")

# Footer
st.divider()
st.markdown("""
### 💡 Tips:
- Paste markdown directly from ChatGPT, Claude, Gemini, or any AI website
- Use GitHub URLs to fetch markdown files from repositories
- Supports headings, lists, bold, italic, code blocks, **LaTeX math**, and **tables**
- Customize font size and styling options in the sidebar
- LaTeX math expressions will be preserved in the document
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
    
    **Supported Features:**
    - Headings (# ## ###)
    - Bold (**text**) and Italic (*text*)
    - Code blocks (```code```)
    - Inline code (`code`)
    - Bullet and numbered lists
    - Tables (| header | header |)
    - LaTeX math: \\( inline \\) and \\[ display \\]
    - Horizontal rules (---)
    
    **GitHub URL Format:**
    - `https://github.com/username/repository/blob/main/file.md`
    - The app will automatically convert it to the raw URL
    """)
