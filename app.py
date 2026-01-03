import streamlit as st
import requests
from io import BytesIO
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re

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
    preserve_math = st.checkbox("Preserve LaTeX math (as green text)", value=True)
    show_debug = st.checkbox("Show debug info", value=False)

# Main content area with tabs
tab1, tab2, tab3 = st.tabs(["📝 Input Markdown", "👁️ Preview", "⬇️ Download"])

with tab1:
    st.subheader("Paste or Edit Your Markdown Content")
    
    # Initialize session state for markdown content
    if 'markdown_content' not in st.session_state:
        st.session_state['markdown_content'] = """# Sample Markdown

## Introduction
This is a **sample** markdown document with *formatting*.

### Math Examples
Inline math: \\( \\alpha \\) and \\( \\beta \\)

Display math:
\\[
S = \\alpha + \\beta
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
    
    def extract_and_format_text(text, paragraph, font_size, preserve_math, debug=False):
        """Extract and format text with inline styles including LaTeX math"""
        
        if debug:
            st.write(f"Processing: {text[:100]}...")
        
        # Process text character by character to handle nested patterns
        i = 0
        current_text = ""
        
        while i < len(text):
            # Check for display math \[ ... \] or \\[ ... \\]
            if preserve_math and i < len(text):
                # Check for single backslash \[
                if text[i:i+2] == '\\[':
                    # Add any accumulated text
                    if current_text:
                        run = paragraph.add_run(current_text)
                        run.font.size = Pt(font_size)
                        current_text = ""
                    
                    # Find closing \]
                    end = text.find('\\]', i + 2)
                    if end != -1:
                        math_content = text[i+2:end].strip()
                        if debug:
                            st.write(f"Found display math (single \\): {math_content}")
                        run = paragraph.add_run(math_content)
                        run.font.name = 'Cambria Math'
                        run.font.size = Pt(font_size)
                        run.font.color.rgb = RGBColor(0, 120, 0)
                        run.bold = True
                        i = end + 2
                        continue
            
            # Check for inline math \( ... \) or \\( ... \\)
            if preserve_math and i < len(text):
                # Check for single backslash \(
                if text[i:i+2] == '\\(':
                    # Add any accumulated text
                    if current_text:
                        run = paragraph.add_run(current_text)
                        run.font.size = Pt(font_size)
                        current_text = ""
                    
                    # Find closing \)
                    end = text.find('\\)', i + 2)
                    if end != -1:
                        math_content = text[i+2:end].strip()
                        if debug:
                            st.write(f"Found inline math (single \\): {math_content}")
                        run = paragraph.add_run(' ' + math_content + ' ')
                        run.font.name = 'Cambria Math'
                        run.font.size = Pt(font_size)
                        run.font.color.rgb = RGBColor(0, 120, 0)
                        run.bold = True
                        i = end + 2
                        continue
            
            # Check for bold **text**
            elif i + 1 < len(text) and text[i:i+2] == '**':
                # Add any accumulated text
                if current_text:
                    run = paragraph.add_run(current_text)
                    run.font.size = Pt(font_size)
                    current_text = ""
                
                # Find closing **
                end = text.find('**', i + 2)
                if end != -1:
                    bold_text = text[i+2:end]
                    run = paragraph.add_run(bold_text)
                    run.bold = True
                    run.font.size = Pt(font_size)
                    i = end + 2
                    continue
            
            # Check for italic *text* (but not **)
            elif text[i] == '*' and (i == 0 or text[i-1] != '*') and (i + 1 >= len(text) or text[i+1] != '*'):
                # Add any accumulated text
                if current_text:
                    run = paragraph.add_run(current_text)
                    run.font.size = Pt(font_size)
                    current_text = ""
                
                # Find closing *
                end = text.find('*', i + 1)
                if end != -1 and (end + 1 >= len(text) or text[end+1] != '*'):
                    italic_text = text[i+1:end]
                    run = paragraph.add_run(italic_text)
                    run.italic = True
                    run.font.size = Pt(font_size)
                    i = end + 1
                    continue
            
            # Check for code `text`
            elif text[i] == '`':
                # Add any accumulated text
                if current_text:
                    run = paragraph.add_run(current_text)
                    run.font.size = Pt(font_size)
                    current_text = ""
                
                # Find closing `
                end = text.find('`', i + 1)
                if end != -1:
                    code_text = text[i+1:end]
                    run = paragraph.add_run(code_text)
                    run.font.name = 'Courier New'
                    run.font.size = Pt(font_size - 1)
                    run.font.color.rgb = RGBColor(220, 50, 50)
                    i = end + 1
                    continue
            
            # Regular character
            current_text += text[i]
            i += 1
        
        # Add any remaining text
        if current_text:
            run = paragraph.add_run(current_text)
            run.font.size = Pt(font_size)
    
    def parse_markdown_to_docx(markdown_text, title, font_size, use_colors, preserve_math, debug=False):
        """Convert markdown to Word document with formatting"""
        doc = Document()
        
        # Add title
        title_para = doc.add_heading(title, 0)
        title_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        
        # Pre-process: handle multi-line display math
        if preserve_math:
            # Find all \[ ... \] blocks (including multi-line)
            display_math_pattern = r'\\\[(.*?)\\\]'
            matches = list(re.finditer(display_math_pattern, markdown_text, re.DOTALL))
            
            # Replace multi-line display math with single line placeholders
            for match in reversed(matches):  # Reverse to maintain positions
                full_match = match.group(0)
                math_content = match.group(1).strip()
                # Replace newlines with spaces in math
                single_line = full_match.replace('\n', ' ')
                markdown_text = markdown_text[:match.start()] + single_line + markdown_text[match.end():]
        
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
                    doc.add_paragraph()
                    table = doc.add_table(rows=1 + len(table_data['rows']), cols=len(table_data['headers']))
                    table.style = 'Light Grid Accent 1'
                    add_table_border(table)
                    
                    # Add headers
                    hdr_cells = table.rows[0].cells
                    for idx, header in enumerate(table_data['headers']):
                        hdr_cells[idx].text = header
                        for p in hdr_cells[idx].paragraphs:
                            for run in p.runs:
                                run.font.bold = True
                                run.font.size = Pt(font_size)
                    
                    # Add rows
                    for row_idx, row_data in enumerate(table_data['rows']):
                        row_cells = table.rows[row_idx + 1].cells
                        for col_idx, cell_data in enumerate(row_data):
                            if col_idx < len(row_cells):
                                row_cells[col_idx].text = cell_data
                                for p in row_cells[col_idx].paragraphs:
                                    for run in p.runs:
                                        run.font.size = Pt(font_size - 1)
                    
                    doc.add_paragraph()
                    i = end_idx
                    continue
            
            # Handle code blocks
            if line.strip().startswith('```'):
                if in_code_block:
                    code_text = '\n'.join(code_lines)
                    para = doc.add_paragraph(code_text)
                    para.style = 'Intense Quote'
                    for run in para.runs:
                        run.font.name = 'Courier New'
                        run.font.size = Pt(font_size - 1)
                    code_lines = []
                    in_code_block = False
                else:
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
                heading_text = line[2:]
                para = doc.add_heading('', 1)
                extract_and_format_text(heading_text, para, font_size + 2, preserve_math, debug)
                if use_colors:
                    for run in para.runs:
                        run.font.color.rgb = RGBColor(0, 51, 102)
            
            elif line.startswith('## ') and not line.startswith('### '):
                heading_text = line[3:]
                para = doc.add_heading('', 2)
                extract_and_format_text(heading_text, para, font_size + 1, preserve_math, debug)
                if use_colors:
                    for run in para.runs:
                        run.font.color.rgb = RGBColor(51, 102, 153)
            
            elif line.startswith('### '):
                heading_text = line[4:]
                para = doc.add_heading('', 3)
                extract_and_format_text(heading_text, para, font_size, preserve_math, debug)
                if use_colors:
                    for run in para.runs:
                        run.font.color.rgb = RGBColor(102, 153, 204)
            
            # Handle bullet lists
            elif line.strip().startswith('- ') or line.strip().startswith('* '):
                text = line.strip()[2:]
                para = doc.add_paragraph(style='List Bullet')
                extract_and_format_text(text, para, font_size, preserve_math, debug)
                in_list = True
            
            # Handle numbered lists
            elif re.match(r'^\d+\.\s', line.strip()):
                text = re.sub(r'^\d+\.\s', '', line.strip())
                para = doc.add_paragraph(style='List Number')
                extract_and_format_text(text, para, font_size, preserve_math, debug)
                in_list = True
            
            # Handle regular paragraphs
            elif line.strip():
                if in_list:
                    doc.add_paragraph()
                    in_list = False
                para = doc.add_paragraph()
                extract_and_format_text(line, para, font_size, preserve_math, debug)
            
            # Handle empty lines
            else:
                if not in_list:
                    doc.add_paragraph()
            
            i += 1
        
        return doc
    
    if st.button("📄 Generate Word Document"):
        with st.spinner("Generating document..."):
            try:
                # Generate the document
                doc = parse_markdown_to_docx(
                    st.session_state['markdown_content'],
                    doc_title,
                    font_size,
                    use_colors,
                    preserve_math,
                    show_debug
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
                import traceback
                st.error(traceback.format_exc())

# Footer
st.divider()
st.markdown("""
### 💡 Tips:
- Paste markdown directly from ChatGPT, Claude, Gemini, or any AI website
- **LaTeX math will appear in GREEN BOLD text** in the Word document
- Math format: `\\( alpha \\)` for inline, `\\[ equation \\]` for display
- Supports headings, lists, bold, italic, code blocks, and tables
- Enable "Show debug info" in sidebar to troubleshoot math rendering
""")

# Instructions section
with st.expander("📖 How to Use"):
    st.markdown("""
    **Method 1: Direct Paste**
    1. Copy markdown content from any AI website (ChatGPT, Claude, etc.)
    2. Paste it into the "Input Markdown" tab
    3. Click "Update Content"
    4. Go to "Download" tab and click "Generate Word Document"
    
    **LaTeX Math Notation:**
    - Inline: `\\( \\alpha \\)` or `\\( x^2 \\)`
    - Display: `\\[ E = mc^2 \\]`
    - Math appears as **green bold text** in Word
    
    **Supported Features:**
    - Headings (# ## ###)
    - Bold (**text**) and Italic (*text*)
    - Code blocks and inline code
    - Bullet and numbered lists
    - Tables
    - LaTeX math expressions
    """)
