import streamlit as st
import requests
from io import BytesIO
from docx import Document
from docx.shared import Pt, RGBColor
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

st.title("📄 Markdown to Word Converter")
st.markdown("✅ **Cloud-ready** - Works on Streamlit Cloud, Heroku, and all platforms!")

# Sidebar
with st.sidebar:
    st.header("🔧 Settings")
    
    st.subheader("GitHub Repository")
    github_url = st.text_input(
        "GitHub File URL",
        placeholder="https://github.com/user/repo/blob/main/file.md"
    )
    
    if st.button("📥 Fetch from GitHub"):
        if github_url:
            try:
                raw_url = github_url.replace("github.com", "raw.githubusercontent.com").replace("/blob/", "/")
                response = requests.get(raw_url, timeout=10)
                response.raise_for_status()
                st.session_state['markdown_content'] = response.text
                st.success("✅ Fetched!")
                st.rerun()
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
        else:
            st.warning("Enter a URL")
    
    st.divider()
    
    st.subheader("Document Settings")
    doc_title = st.text_input("Title", value="Converted Document")
    font_size = st.slider("Font Size", 8, 16, 11)
    use_colors = st.checkbox("Colored headings", value=True)
    
    st.divider()
    st.success("✅ Ready for Streamlit Cloud!")

# Initialize
if 'markdown_content' not in st.session_state:
    st.session_state['markdown_content'] = """# Sample Document

## Introduction
This is a **sample** with *formatting*.

### Math Examples
Inline: \\( \\alpha \\) and \\( \\beta \\)

Display:
\\[
E = mc^2
\\]

### Features
- Bullet 1
- Bullet 2

| Column 1 | Column 2 |
|----------|----------|
| Data 1   | Data 2   |
"""

# Helper functions
def add_table_border(table):
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
    table_lines = []
    idx = start_idx
    while idx < len(lines) and '|' in lines[idx]:
        table_lines.append(lines[idx])
        idx += 1
    if len(table_lines) < 2:
        return None, start_idx
    headers = [cell.strip() for cell in table_lines[0].split('|') if cell.strip()]
    rows = []
    for line in table_lines[2:]:
        cells = [cell.strip() for cell in line.split('|') if cell.strip()]
        if cells:
            rows.append(cells)
    return {'headers': headers, 'rows': rows}, idx

def convert_latex_to_unicode(latex_text):
    result = latex_text
    result = result.replace(r'\text{', '').replace(r'\mathrm{', '').replace(r'\mathbf{', '')
    result = result.replace(r'\hat{', '').replace(r'\bar{', '').replace(r'\tilde{', '')
    result = result.replace(r'\boxed{', '').replace(r'\left', '').replace(r'\right', '')
    result = result.replace(r'\begin{aligned}', '').replace(r'\end{aligned}', '')
    result = result.replace(r'\,', ' ').replace(r'\quad', '  ')
    
    replacements = {
        r'\alpha': 'α', r'\beta': 'β', r'\gamma': 'γ', r'\delta': 'δ',
        r'\epsilon': 'ε', r'\theta': 'θ', r'\lambda': 'λ', r'\mu': 'μ',
        r'\nu': 'ν', r'\xi': 'ξ', r'\pi': 'π', r'\rho': 'ρ', r'\sigma': 'σ',
        r'\tau': 'τ', r'\phi': 'φ', r'\chi': 'χ', r'\psi': 'ψ', r'\omega': 'ω',
        r'\Gamma': 'Γ', r'\Delta': 'Δ', r'\Theta': 'Θ', r'\Lambda': 'Λ',
        r'\Pi': 'Π', r'\Sigma': 'Σ', r'\Phi': 'Φ', r'\Psi': 'Ψ', r'\Omega': 'Ω',
        r'\times': '×', r'\div': '÷', r'\pm': '±', r'\cdot': '·',
        r'\leq': '≤', r'\geq': '≥', r'\neq': '≠', r'\approx': '≈', r'\equiv': '≡',
        r'\rightarrow': '→', r'\to': '→', r'\leftarrow': '←', r'\leftrightarrow': '↔',
        r'\in': '∈', r'\infty': '∞', r'\int': '∫', r'\sum': '∑', r'\partial': '∂',
        r'\hbar': 'ℏ', r'\sqrt': '√',
    }
    
    for latex, unicode_char in replacements.items():
        result = result.replace(latex, unicode_char)
    
    result = re.sub(r'\\frac\{([^}]+)\}\{([^}]+)\}', r'(\1)/(\2)', result)
    
    def to_super(m):
        s = {'0':'⁰','1':'¹','2':'²','3':'³','4':'⁴','5':'⁵','6':'⁶','7':'⁷','8':'⁸','9':'⁹','+':'⁺','-':'⁻'}
        return ''.join(s.get(c,c) for c in m.group(1))
    
    def to_sub(m):
        s = {'0':'₀','1':'₁','2':'₂','3':'₃','4':'₄','5':'₅','6':'₆','7':'₇','8':'₈','9':'₉','s':'ₛ','n':'ₙ','z':'ᵤ'}
        return ''.join(s.get(c,c) for c in m.group(1))
    
    result = re.sub(r'\^\{([^}]+)\}', to_super, result)
    result = re.sub(r'_\{([^}]+)\}', to_sub, result)
    result = result.replace('{', '').replace('}', '')
    result = re.sub(r'\\[a-zA-Z]+', '', result)
    result = result.replace('\\', '')
    return result.strip()

def format_text(text, para, size):
    i = 0
    curr = ""
    
    while i < len(text):
        done = False
        
        if text[i:i+2] == '\\[':
            if curr:
                para.add_run(curr).font.size = Pt(size)
                curr = ""
            end = text.find('\\]', i+2)
            if end != -1:
                math = convert_latex_to_unicode(text[i+2:end].strip())
                r = para.add_run(math)
                r.font.name = 'Cambria Math'
                r.font.size = Pt(size+1)
                r.font.color.rgb = RGBColor(0,120,0)
                r.bold = True
                i = end+2
                done = True
        
        if not done and text[i:i+2] == '\\(':
            if curr:
                para.add_run(curr).font.size = Pt(size)
                curr = ""
            end = text.find('\\)', i+2)
            if end != -1:
                math = convert_latex_to_unicode(text[i+2:end].strip())
                r = para.add_run(' '+math+' ')
                r.font.name = 'Cambria Math'
                r.font.size = Pt(size)
                r.font.color.rgb = RGBColor(0,120,0)
                r.bold = True
                i = end+2
                done = True
        
        if not done and text[i:i+2] == '**':
            if curr:
                para.add_run(curr).font.size = Pt(size)
                curr = ""
            end = text.find('**', i+2)
            if end != -1:
                r = para.add_run(text[i+2:end])
                r.bold = True
                r.font.size = Pt(size)
                i = end+2
                done = True
        
        if not done and text[i] == '*' and (i==0 or text[i-1]!='*') and (i+1>=len(text) or text[i+1]!='*'):
            if curr:
                para.add_run(curr).font.size = Pt(size)
                curr = ""
            end = i+1
            while end < len(text) and not (text[end]=='*' and (end+1>=len(text) or text[end+1]!='*')):
                end += 1
            if end < len(text):
                r = para.add_run(text[i+1:end])
                r.italic = True
                r.font.size = Pt(size)
                i = end+1
                done = True
        
        if not done and text[i] == '`':
            if curr:
                para.add_run(curr).font.size = Pt(size)
                curr = ""
            end = text.find('`', i+1)
            if end != -1:
                r = para.add_run(text[i+1:end])
                r.font.name = 'Courier New'
                r.font.size = Pt(size-1)
                r.font.color.rgb = RGBColor(220,50,50)
                i = end+1
                done = True
        
        if not done:
            curr += text[i]
            i += 1
    
    if curr:
        para.add_run(curr).font.size = Pt(size)

def convert_to_docx(md_text, title, size, colors):
    doc = Document()
    
    tp = doc.add_heading(title, 0)
    tp.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    
    if '\\[' in md_text:
        md_text = re.sub(r'\\\[(.*?)\\\]', lambda m: '\\['+m.group(1).replace('\n',' ')+'\\]', md_text, flags=re.DOTALL)
    
    lines = md_text.split('\n')
    in_code = False
    code = []
    i = 0
    
    while i < len(lines):
        line = lines[i]
        
        if '|' in line and i+1<len(lines) and '---' in lines[i+1]:
            td, ei = parse_table(lines, i)
            if td:
                t = doc.add_table(rows=1+len(td['rows']), cols=len(td['headers']))
                t.style = 'Light Grid Accent 1'
                add_table_border(t)
                for idx, h in enumerate(td['headers']):
                    t.rows[0].cells[idx].text = h
                    for p in t.rows[0].cells[idx].paragraphs:
                        for r in p.runs:
                            r.font.bold = True
                            r.font.size = Pt(size)
                for ri, row in enumerate(td['rows']):
                    for ci, cell in enumerate(row):
                        if ci < len(t.rows[ri+1].cells):
                            t.rows[ri+1].cells[ci].text = cell
                i = ei
                continue
        
        if line.strip().startswith('```'):
            if in_code:
                p = doc.add_paragraph('\n'.join(code))
                p.style = 'Intense Quote'
                for r in p.runs:
                    r.font.name = 'Courier New'
                    r.font.size = Pt(size-1)
                code = []
                in_code = False
            else:
                in_code = True
            i += 1
            continue
        
        if in_code:
            code.append(line)
            i += 1
            continue
        
        if line.strip() == '---':
            p = doc.add_paragraph('─'*80)
            for r in p.runs:
                r.font.color.rgb = RGBColor(200,200,200)
            i += 1
            continue
        
        if line.startswith('# ') and not line.startswith('## '):
            p = doc.add_heading('', 1)
            format_text(line[2:], p, size+2)
            if colors:
                for r in p.runs:
                    if not r.font.color.rgb:
                        r.font.color.rgb = RGBColor(0,51,102)
        elif line.startswith('## ') and not line.startswith('### '):
            p = doc.add_heading('', 2)
            format_text(line[3:], p, size+1)
            if colors:
                for r in p.runs:
                    if not r.font.color.rgb:
                        r.font.color.rgb = RGBColor(51,102,153)
        elif line.startswith('### '):
            p = doc.add_heading('', 3)
            format_text(line[4:], p, size)
            if colors:
                for r in p.runs:
                    if not r.font.color.rgb:
                        r.font.color.rgb = RGBColor(102,153,204)
        elif line.strip().startswith('- ') or line.strip().startswith('* '):
            p = doc.add_paragraph(style='List Bullet')
            format_text(line.strip()[2:], p, size)
        elif re.match(r'^\d+\.\s', line.strip()):
            p = doc.add_paragraph(style='List Number')
            format_text(re.sub(r'^\d+\.\s','',line.strip()), p, size)
        elif line.strip().startswith('>'):
            p = doc.add_paragraph()
            p.style = 'Intense Quote'
            format_text(line.strip()[1:].strip(), p, size)
        elif line.strip():
            p = doc.add_paragraph()
            format_text(line, p, size)
        
        i += 1
    
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio.getvalue()

# Main UI
st.subheader("📝 Paste Markdown")

st.info("**Math:** Use `\\(...\\)` for inline and `\\[...\\]` for display")

md_input = st.text_area(
    "Content",
    value=st.session_state['markdown_content'],
    height=400
)

st.session_state['markdown_content'] = md_input

if st.button("🔄 Convert to Word", type="primary", use_container_width=True):
    with st.spinner("Processing..."):
        try:
            docx_data = convert_to_docx(
                st.session_state['markdown_content'],
                doc_title,
                font_size,
                use_colors
            )
            
            st.success("✅ Ready!")
            
            st.download_button(
                "⬇️ Download Word",
                data=docx_data,
                file_name=f"{doc_title.replace(' ','_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True
            )
            
        except Exception as e:
            st.error(f"❌ {str(e)}")

st.divider()
st.markdown("""
### ✅ Cloud Deployment Ready

**requirements.txt:**
```
streamlit
python-docx
requests
```

**Deploy to Streamlit Cloud:**
1. Push code to GitHub
2. Go to share.streamlit.io
3. Deploy!

**Features:**
- ✅ Works on all cloud platforms
- ✅ No external dependencies
- ✅ LaTeX → Unicode (α, β, etc.)
- ✅ Tables, lists, code blocks
- ✅ GitHub import
""")
