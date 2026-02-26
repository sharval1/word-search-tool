"""
Word Document Search Tool
Drop .docx files and search by keyword to get all matching content.
"""
import html
import re
import streamlit as st
from search_engine import (
    extract_text_from_docx,
    extract_images_from_docx,
    get_nearest_image,
    search_keyword,
)

st.set_page_config(page_title="Word Search Tool", page_icon="📄", layout="centered")
st.title("📄 Word Document Search")
st.caption("Upload your Word files, then search by keyword to find all related content.")

# File uploader - drop multiple Word files
uploaded_files = st.file_uploader(
    "Drop your Word files here",
    type=["docx"],
    accept_multiple_files=True,
    help="Select one or more .docx files",
)

# Store extracted content in session so we don't re-parse on every keystroke
if "all_paragraphs" not in st.session_state:
    st.session_state.all_paragraphs = []
if "all_images" not in st.session_state:
    st.session_state.all_images = []
if "file_bytes_by_name" not in st.session_state:
    st.session_state.file_bytes_by_name = {}  # doc_name -> file bytes

if uploaded_files:
    file_ids = [f.name + str(f.size) for f in uploaded_files]
    if "last_file_ids" not in st.session_state or st.session_state.last_file_ids != file_ids:
        st.session_state.last_file_ids = file_ids
        all_paragraphs = []
        all_images = []
        file_bytes_by_name = {}
        for f in uploaded_files:
            b = f.getvalue()
            file_bytes_by_name[f.name] = b
            all_paragraphs.extend(extract_text_from_docx(b, f.name))
            all_images.extend(extract_images_from_docx(b, f.name))
        st.session_state.all_paragraphs = all_paragraphs
        st.session_state.all_images = all_images
        st.session_state.file_bytes_by_name = file_bytes_by_name

    total_images = len(st.session_state.all_images)
    st.success(f"Loaded **{len(uploaded_files)}** file(s) • **{len(st.session_state.all_paragraphs)}** text blocks • **{total_images}** image(s) indexed.")

    # Search box
    keyword = st.text_input(
        "Search keyword",
        placeholder="Type any word or phrase...",
        key="keyword",
    )

    if keyword:
        matches = search_keyword(st.session_state.all_paragraphs, keyword)
        if matches:
            docs_with_matches = set(fname for fname, _ in matches)
            all_paragraphs = st.session_state.all_paragraphs
            all_images = st.session_state.all_images
            images_by_doc = {}
            for doc_name, img_bytes, img_name in all_images:
                images_by_doc.setdefault(doc_name, []).append((img_bytes, img_name))

            st.markdown(f"**{len(docs_with_matches)}** document(s) found for **\"{keyword}\"**.")

            file_bytes_by_name = st.session_state.get("file_bytes_by_name", {})

            for doc_name in sorted(docs_with_matches):
                doc_paras = [(f, t) for f, t in all_paragraphs if f == doc_name]
                indices_to_include = set()
                matching_texts_for_doc = []
                for (f, t) in matches:
                    if f != doc_name:
                        continue
                    matching_texts_for_doc.append(t)
                    for i, (_, pt) in enumerate(doc_paras):
                        if pt == t:
                            indices_to_include.add(i)
                            if i > 0:
                                indices_to_include.add(i - 1)
                            if i < len(doc_paras) - 1:
                                indices_to_include.add(i + 1)
                            break
                sorted_indices = sorted(indices_to_include)
                complete_text = "\n\n".join(doc_paras[i][1] for i in sorted_indices)

                def highlight(s):
                    return re.sub(re.escape(keyword), lambda m: f"<strong>{m.group(0)}</strong>", s, flags=re.IGNORECASE)
                highlighted = highlight(html.escape(complete_text)).replace("\n", "<br>")
                doc_imgs = images_by_doc.get(doc_name, [])
                # Pick image nearest to keyword in document (or first if parsing fails)
                nearest = None
                if doc_imgs and doc_name in file_bytes_by_name:
                    all_imgs_for_doc = [(doc_name, b, n) for _dn, b, n in all_images if _dn == doc_name]
                    nearest = get_nearest_image(
                        file_bytes_by_name[doc_name],
                        doc_name,
                        matching_texts_for_doc,
                        all_imgs_for_doc,
                    )
                if not nearest and doc_imgs:
                    nearest = (doc_imgs[0][0], doc_imgs[0][1])

                with st.container():
                    st.markdown("---")
                    st.markdown(f"### 📄 {doc_name}")
                    # Layout: text and image side by side (image on right)
                    col_text, col_img = st.columns([2, 1])
                    with col_text:
                        st.markdown("**Related information**")
                        st.markdown(
                            f'<div style="background-color:#f8f9fa; padding:14px; border-radius:8px; border-left:4px solid #1f77b4;">{highlighted}</div>',
                            unsafe_allow_html=True,
                        )
                    with col_img:
                        if nearest:
                            st.markdown("**Relevant image**")
                            img_bytes, img_name = nearest
                            try:
                                st.image(img_bytes, caption=img_name, use_container_width=True)
                            except Exception:
                                st.caption(f"*{img_name}*")
                    st.markdown("---")
        else:
            st.info(f"No results found for **\"{keyword}\"** in the uploaded documents.")
else:
    st.info("👆 Upload one or more Word (.docx) files to start searching.")
