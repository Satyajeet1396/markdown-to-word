import streamlit as st
import pypandoc
import tempfile
import os

# ======================================================
# Ensure Pandoc is available (CRITICAL for Streamlit Cloud)
# ======================================================
try:
    pandoc_version = pypandoc.get_pandoc_version()
except OSError:
    pypandoc.download_pandoc()
    pandoc_version = pypandoc.get_pandoc_version()

# ======================================================
# Streamlit UI
# ======================================================
st.set_page_config(
    page_title="Markdown to Word Converter",
    layout="centered"
)

st.title("📄 Markdown → Word (.docx) Converter")
st.caption(f"Pandoc version: {pandoc_version}")

st.markdown(
    """
✅ Supports:
- Headings  
- Tables  
- Lists  
- Images  
- **LaTeX equations** (`$...$`, `$$...$$`)  
"""
)

md_text = st.text_area(
    "✍️ Paste your Markdown content here:",
    height=300,
    placeholder="# Sample\n\nThis is **Markdown** with $E=mc^2$"
)

uploaded_md = st.file_uploader(
    "📂 Or upload a Markdown (.md) file",
    type=["md"]
)

convert_btn = st.button("🚀 Convert to Word")

# ======================================================
# Conversion Logic
# ======================================================
if convert_btn:
    if not md_text and not uploaded_md:
        st.warning("⚠️ Please paste Markdown or upload a .md file.")
    else:
        with tempfile.TemporaryDirectory() as tmpdir:
            md_path = os.path.join(tmpdir, "input.md")
            docx_path = os.path.join(tmpdir, "output.docx")

            if uploaded_md:
                md_text = uploaded_md.read().decode("utf-8")

            with open(md_path, "w", encoding="utf-8") as f:
                f.write(md_text)

            try:
                pypandoc.convert_file(
                    md_path,
                    "docx",
                    outputfile=docx_path,
                    extra_args=["--standalone"]
                )

                with open(docx_path, "rb") as f:
                    st.success("✅ Conversion successful!")
                    st.download_button(
                        label="⬇️ Download Word File",
                        data=f,
                        file_name="converted.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

            except Exception as e:
                st.error("❌ Conversion failed")
                st.exception(e)

# ======================================================
# Footer
# ======================================================
st.markdown("---")
st.caption("🔬 Designed for researchers, students & academic writing")
