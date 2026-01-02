import streamlit as st
import pypandoc
import tempfile
import os
from pathlib import Path

# ======================================================
# FORCE Pandoc installation & path registration
# ======================================================
PANDOC_DIR = Path.home() / ".pandoc"
PANDOC_BIN = PANDOC_DIR / "pandoc"

def ensure_pandoc():
    try:
        return pypandoc.get_pandoc_version()
    except OSError:
        pypandoc.download_pandoc(targetfolder=str(PANDOC_DIR))
        os.environ["PYPANDOC_PANDOC"] = str(PANDOC_BIN)
        os.environ["PATH"] += os.pathsep + str(PANDOC_DIR)
        return pypandoc.get_pandoc_version()

pandoc_version = ensure_pandoc()

# ======================================================
# Streamlit UI
# ======================================================
st.set_page_config(page_title="Markdown → Word Converter")

st.title("📄 Markdown → Word (.docx) Converter")
st.caption(f"Pandoc version: {pandoc_version}")

st.markdown("""
✅ Supports:
- Headings
- Tables
- Lists
- Images
- **LaTeX equations** (`$...$`, `$$...$$`)
""")

md_text = st.text_area(
    "✍️ Paste Markdown here",
    height=300,
    placeholder="# Example\n\nEquation: $E = mc^2$"
)

uploaded_md = st.file_uploader(
    "📂 Or upload a .md file",
    type=["md"]
)

# ======================================================
# Conversion
# ======================================================
if st.button("🚀 Convert to Word"):
    if not md_text and not uploaded_md:
        st.warning("⚠️ Please provide Markdown input.")
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
                    source_file=md_path,
                    to="docx",
                    outputfile=docx_path,
                    extra_args=["--standalone"]
                )

                with open(docx_path, "rb") as f:
                    st.success("✅ Conversion successful")
                    st.download_button(
                        "⬇️ Download Word File",
                        f,
                        file_name="converted.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

            except Exception as e:
                st.error("❌ Conversion failed")
                st.exception(e)

st.markdown("---")
st.caption("🔬 Academic-grade Markdown → Word converter")
