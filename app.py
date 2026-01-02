import streamlit as st
import tempfile
import subprocess
import os
import shutil

# ======================================================
# Check if pandoc exists (Streamlit Cloud has it)
# ======================================================
pandoc_path = shutil.which("pandoc")

if pandoc_path is None:
    st.error("❌ Pandoc is not available in this environment.")
    st.stop()

# ======================================================
# UI
# ======================================================
st.set_page_config(page_title="Markdown → Word Converter")
st.title("📄 Markdown → Word (.docx) Converter")
st.caption(f"Using Pandoc: {pandoc_path}")

md_text = st.text_area(
    "✍️ Paste Markdown here",
    height=300,
    placeholder="# Example\n\nEquation:\n$$E = mc^2$$"
)

uploaded_md = st.file_uploader(
    "📂 Or upload a Markdown (.md) file",
    type=["md"]
)

# ======================================================
# Convert
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
                subprocess.run(
                    [
                        pandoc_path,
                        md_path,
                        "-o",
                        docx_path,
                        "--standalone"
                    ],
                    check=True
                )

                with open(docx_path, "rb") as f:
                    st.success("✅ Conversion successful")
                    st.download_button(
                        "⬇️ Download Word File",
                        f,
                        file_name="converted.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

            except subprocess.CalledProcessError as e:
                st.error("❌ Pandoc conversion failed")
                st.code(str(e))
