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
    
    def convert_latex_to_unicode(latex_text):
        """Convert common LaTeX symbols to Unicode characters"""
        # First, remove LaTeX formatting commands
        result = latex_text
        
        # Remove common LaTeX commands that don't affect display
        result = result.replace(r'\text{', '').replace(r'\mathrm{', '').replace(r'\mathbf{', '')
        result = result.replace(r'\hat{', '').replace(r'\bar{', '').replace(r'\tilde{', '')
        result = result.replace(r'\boxed{', '').replace(r'\left', '').replace(r'\right', '')
        result = result.replace(r'\begin{aligned}', '').replace(r'\end{aligned}', '')
        result = result.replace(r'\,', ' ').replace(r'\;', ' ').replace(r'\:', ' ')
        result = result.replace(r'\quad', '  ').replace(r'\qquad', '    ')
        
        # Remove extra closing braces
        while result.count('}') > result.count('{'):
            result = result.replace('}', '', 1)
        
        replacements = {
            # Greek letters (lowercase)
            r'\alpha': 'α', r'\beta': 'β', r'\gamma': 'γ', r'\delta': 'δ',
            r'\epsilon': 'ε', r'\varepsilon': 'ε', r'\zeta': 'ζ', r'\eta': 'η', 
            r'\theta': 'θ', r'\vartheta': 'θ', r'\iota': 'ι', r'\kappa': 'κ', 
            r'\lambda': 'λ', r'\mu': 'μ', r'\nu': 'ν', r'\xi': 'ξ', 
            r'\pi': 'π', r'\rho': 'ρ', r'\sigma': 'σ', r'\varsigma': 'ς',
            r'\tau': 'τ', r'\upsilon': 'υ', r'\phi': 'φ', r'\varphi': 'φ',
            r'\chi': 'χ', r'\psi': 'ψ', r'\omega': 'ω',
            # Greek letters (uppercase)
            r'\Alpha': 'Α', r'\Beta': 'Β', r'\Gamma': 'Γ', r'\Delta': 'Δ',
            r'\Epsilon': 'Ε', r'\Zeta': 'Ζ', r'\Eta': 'Η', r'\Theta': 'Θ',
            r'\Iota': 'Ι', r'\Kappa': 'Κ', r'\Lambda': 'Λ', r'\Mu': 'Μ',
            r'\Nu': 'Ν', r'\Xi': 'Ξ', r'\Pi': 'Π', r'\Rho': 'Ρ',
            r'\Sigma': 'Σ', r'\Tau': 'Τ', r'\Upsilon': 'Υ', r'\Phi': 'Φ',
            r'\Chi': 'Χ', r'\Psi': 'Ψ', r'\Omega': 'Ω',
            # Math operators
            r'\times': '×', r'\div': '÷', r'\pm': '±', r'\mp': '∓',
            r'\cdot': '·', r'\ast': '∗', r'\star': '⋆',
            # Relations
            r'\leq': '≤', r'\geq': '≥', r'\neq': '≠', r'\ne': '≠',
            r'\approx': '≈', r'\equiv': '≡', r'\sim': '∼', r'\simeq': '≃',
            r'\propto': '∝', r'\ll': '≪', r'\gg': '≫',
            # Arrows
            r'\rightarrow': '→', r'\to': '→', r'\leftarrow': '←', 
            r'\leftrightarrow': '↔', r'\Rightarrow': '⇒', 
            r'\Leftarrow': '⇐', r'\Leftrightarrow': '⇔',
            r'\uparrow': '↑', r'\downarrow': '↓',
            # Sets
            r'\in': '∈', r'\notin': '∉', r'\ni': '∋',
            r'\subset': '⊂', r'\supset': '⊃', r'\subseteq': '⊆', r'\supseteq': '⊇',
            r'\cup': '∪', r'\cap': '∩', r'\emptyset': '∅', r'\varnothing': '∅',
            r'\infty': '∞', r'\forall': '∀', r'\exists': '∃',
            # Calculus
            r'\int': '∫', r'\iint': '∬', r'\iiint': '∭', r'\oint': '∮',
            r'\sum': '∑', r'\prod': '∏',
            r'\partial': '∂', r'\nabla': '∇',
            # Other symbols
            r'\hbar': 'ℏ', r'\ell': 'ℓ', r'\wp': '℘',
            r'\Re': 'ℜ', r'\Im': 'ℑ',
            r'\aleph': 'ℵ', r'\beth': 'ℶ',
            r'\sqrt': '√', r'\angle': '∠', r'\degree': '°',
            r'\circ': '∘', r'\bullet': '•',
            r'\langle': '⟨', r'\rangle': '⟩',
            # Special brackets
            r'\{': '{', r'\}': '}',
        }
        
        for latex, unicode_char in replacements.items():
            result = result.replace(latex, unicode_char)
        
        # Handle sqrt with braces: \sqrt{2} -> √2
        result = re.sub(r'√\{([^}]+)\}', r'√(\1)', result)
        
        # Handle fractions \frac{a}{b} -> (a)/(b)
        frac_pattern = r'\\frac\{([^}]+)\}\{([^}]+)\}'
        result = re.sub(frac_pattern, r'(\1)/(\2)', result)
        
        # Handle superscripts with braces: x^{2} -> x²
        def convert_superscript(match):
            text = match.group(1)
            superscript_map = {
                '0': '⁰', '1': '¹', '2': '²', '3': '³', '4': '⁴',
                '5': '⁵', '6': '⁶', '7': '⁷', '8': '⁸', '9': '⁹',
                '+': '⁺', '-': '⁻', '=': '⁼', '(': '⁽', ')': '⁾',
                'n': 'ⁿ', 'i': 'ⁱ'
            }
            return ''.join(superscript_map.get(c, c) for c in text)
        
        result = re.sub(r'\^\{([^}]+)\}', convert_superscript, result)
        result = re.sub(r'\^(\d)', convert_superscript, result)
        
        # Handle subscripts with braces: x_{1} -> x₁
        def convert_subscript(match):
            text = match.group(1)
            subscript_map = {
                '0': '₀', '1': '₁', '2': '₂', '3': '₃', '4': '₄',
                '5': '₅', '6': '₆', '7': '₇', '8': '₈', '9': '₉',
                '+': '₊', '-': '₋', '=': '₌', '(': '₍', ')': '₎',
                'a': 'ₐ', 'e': 'ₑ', 'o': 'ₒ', 'x': 'ₓ', 'h': 'ₕ',
                'k': 'ₖ', 'l': 'ₗ', 'm': 'ₘ', 'n': 'ₙ', 'p': 'ₚ',
                's': 'ₛ', 't': 'ₜ'
            }
            return ''.join(subscript_map.get(c, c) for c in text)
        
        result = re.sub(r'_\{([^}]+)\}', convert_subscript, result)
        result = re.sub(r'_(\d)', convert_subscript, result)
        
        # Clean up remaining single braces
        result = result.replace('{', '').replace('}', '')
        
        # Clean up backslashes for commands we might have missed
        result = re.sub(r'\\[a-zA-Z]+', '', result)
        result = result.replace('\\', '')
        
        return result.strip()
    
    def extract_and_format_text(text, paragraph, font_size, preserve_math, debug=False):
        """Extract and format text with inline styles including LaTeX math"""
        
        if debug:
            st.write(f"Processing: {text[:100]}...")
        
        # Process text character by character to handle all patterns
        i = 0
        current_text = ""
        
        while i < len(text):
            processed = False
            
            # Check for display math \[ ... \]
            if preserve_math and text[i:i+2] == '\\[':
                if current_text:
                    run = paragraph.add_run(current_text)
                    run.font.size = Pt(font_size)
                    current_text = ""
                
                end = text.find('\\]', i + 2)
                if end != -1:
                    math_content = text[i+2:end].strip()
                    # Convert LaTeX to Unicode
                    unicode_math = convert_latex_to_unicode(math_content)
                    if debug:
                        st.write(f"Found display math: {math_content} → {unicode_math}")
                    run = paragraph.add_run('\n' + unicode_math + '\n')
                    run.font.name = 'Cambria Math'
                    run.font.size = Pt(font_size + 1)
                    run.font.color.rgb = RGBColor(0, 120, 0)
                    run.bold = True
                    i = end + 2
                    processed = True
            
            # Check for inline math \( ... \)
            if not processed and preserve_math and text[i:i+2] == '\\(':
                if current_text:
                    run = paragraph.add_run(current_text)
                    run.font.size = Pt(font_size)
                    current_text = ""
                
                end = text.find('\\)', i + 2)
                if end != -1:
                    math_content = text[i+2:end].strip()
                    # Convert LaTeX to Unicode
                    unicode_math = convert_latex_to_unicode(math_content)
                    if debug:
                        st.write(f"Found inline math: {math_content} → {unicode_math}")
                    run = paragraph.add_run(' ' + unicode_math + ' ')
                    run.font.name = 'Cambria Math'
                    run.font.size = Pt(font_size)
                    run.font.color.rgb = RGBColor(0, 120, 0)
                    run.bold = True
                    i = end + 2
                    processed = True
            
            # Check for bold **text**
            if not processed and text[i:i+2] == '**':
                if current_text:
                    run = paragraph.add_run(current_text)
                    run.font.size = Pt(font_size)
                    current_text = ""
                
                end = text.find('**', i + 2)
                if end != -1 and end > i + 2:
                    bold_text = text[i+2:end]
                    run = paragraph.add_run(bold_text)
                    run.bold = True
                    run.font.size = Pt(font_size)
                    i = end + 2
                    processed = True
            
            # Check for italic *text* (single asterisk, not part of **)
            if not processed and text[i] == '*':
                # Make sure it's not part of **
                if (i == 0 or text[i-1] != '*') and (i + 1 >= len(text) or text[i+1] != '*'):
                    if current_text:
                        run = paragraph.add_run(current_text)
                        run.font.size = Pt(font_size)
                        current_text = ""
                    
                    # Find next single *
                    end = i + 1
                    while end < len(text):
                        if text[end] == '*' and (end + 1 >= len(text) or text[end+1] != '*') and (end == 0 or text[end-1] != '*'):
                            italic_text = text[i+1:end]
                            if italic_text:  # Only if there's content
                                run = paragraph.add_run(italic_text)
                                run.italic = True
                                run.font.size = Pt(font_size)
                                i = end + 1
                                processed = True
                            break
                        end += 1
            
            # Check for inline code `text`
            if not processed and text[i] == '`':
                if current_text:
                    run = paragraph.add_run(current_text)
                    run.font.size = Pt(font_size)
                    current_text = ""
                
                end = text.find('`', i + 1)
                if end != -1:
                    code_text = text[i+1:end]
                    run = paragraph.add_run(code_text)
                    run.font.name = 'Courier New'
                    run.font.size = Pt(font_size - 1)
                    run.font.color.rgb = RGBColor(220, 50, 50)
                    i = end + 1
                    processed = True
            
            # If nothing was processed, add character to current_text
            if not processed:
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
- **LaTeX math will be converted to Unicode symbols** (α, β, →, ∫, etc.) and appear in **GREEN BOLD**
- Math format: `\( \alpha \)` for inline → displays as green **α**
- Display math: `\[ E = mc^2 \]` → displays as green equation
- Supports Greek letters, operators, arrows, calculus symbols, and more
- Enable "Show debug info" to see LaTeX → Unicode conversion
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
