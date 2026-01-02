import streamlit as st
import pypandoc
import tempfile
import os

st.set_page_config(page_title="Markdown → Word Converter")

st.title("📄 Markdown to Word (.docx) Converter")

md_text = st.text_area("Paste your Markdown here", height=300)

uploaded_md = st.file_uploader("Or upload a .md file", type=["md"])

if st.button("Convert to Word"):
    if md_text or uploaded_md:
        with tempfile.TemporaryDirectory() as tmpdir:
            md_path = os.path.join(tmpdir, "input.md")
            docx_path = os.path.join(tmpdir, "output.docx")

            if uploaded_md:
                md_text = uploaded_md.read().decode("utf-8")

            with open(md_path, "w", encoding="utf-8") as f:
                f.write(md_text)

            pypandoc.convert_file(
                md_path,
                "docx",
                outputfile=docx_path,
                extra_args=["--standalone"]
            )

            with open(docx_path, "rb") as f:
                st.download_button(
                    "⬇️ Download Word File",
                    f,
                    file_name="converted.docx"
                )
    else:
        st.warning("Please provide Markdown input.")
