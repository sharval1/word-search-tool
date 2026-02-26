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

st.set_page_config(page_title="Word Search Tool", page_icon="📄", layout="wide")
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

            file_bytes_by_name = st.session_state.get("file_bytes_by_name", {})

            # Build result data for each document (for sorting + expanders)
            def highlight(s):
                return re.sub(re.escape(keyword), lambda m: f"<strong>{m.group(0)}</strong>", s, flags=re.IGNORECASE)

            doc_results = []
            for doc_name in docs_with_matches:
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
                doc_keyword_count = len(re.findall(re.escape(keyword), complete_text, re.IGNORECASE))
                highlighted = highlight(html.escape(complete_text)).replace("\n", "<br>")
                doc_imgs = images_by_doc.get(doc_name, [])
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
                doc_results.append({
                    "name": doc_name,
                    "count": doc_keyword_count,
                    "complete_text": complete_text,
                    "highlighted": highlighted,
                    "nearest": nearest,
                })

            total_count = sum(r["count"] for r in doc_results)

            # ---- Sidebar: interactive options ----
            with st.sidebar:
                st.markdown("### ⚙️ Options")
                sort_by = st.radio(
                    "Sort results by",
                    ["Document name (A–Z)", "Keyword count (highest first)"],
                    key="sort_by",
                )
                show_images = st.checkbox("Show relevant image", value=True, key="show_images")
                min_count = st.slider(
                    "Minimum keyword count (filter documents)",
                    min_value=0,
                    max_value=max((r["count"] for r in doc_results), default=0),
                    value=0,
                    key="min_count",
                )
                st.markdown("---")
                # Download summary
                summary_lines = [
                    f"Keyword: \"{keyword}\"",
                    f"Total occurrences: {total_count}",
                    f"Documents: {len(doc_results)}",
                    "",
                ]
                for r in doc_results:
                    if r["count"] >= min_count:
                        summary_lines.append(f"  - {r['name']}: {r['count']} time(s)")
                summary_text = "\n".join(summary_lines)
                st.download_button(
                    "📥 Download results summary",
                    data=summary_text,
                    file_name=f"search_results_{keyword[:20].replace(' ', '_')}.txt",
                    mime="text/plain",
                    key="download_summary",
                )

            # Filter by min count
            doc_results = [r for r in doc_results if r["count"] >= min_count]
            # Sort
            if sort_by == "Keyword count (highest first)":
                doc_results = sorted(doc_results, key=lambda r: (-r["count"], r["name"]))
            else:
                doc_results = sorted(doc_results, key=lambda r: r["name"])

            # ---- Main area: metrics + expandable results ----
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Documents found", len(doc_results))
            with col2:
                st.metric("Total keyword count", total_count)
            with col3:
                st.metric("Keyword", f"\"{keyword}\"")
            st.markdown("---")

            for r in doc_results:
                with st.expander(f"📄 **{r['name']}** — keyword appears **{r['count']}** time(s)", expanded=True):
                    col_text, col_img = st.columns([2, 1])
                    with col_text:
                        st.markdown("**Related information**")
                        st.markdown(
                            f'<div style="background-color:#f8f9fa; padding:14px; border-radius:8px; border-left:4px solid #1f77b4;">{r["highlighted"]}</div>',
                            unsafe_allow_html=True,
                        )
                    with col_img:
                        if show_images and r["nearest"]:
                            st.markdown("**Relevant image**")
                            img_bytes, img_name = r["nearest"]
                            try:
                                st.image(img_bytes, caption=img_name, use_container_width=True)
                            except Exception:
                                st.caption(f"*{img_name}*")
            st.markdown("---")
        else:
            st.info(f"No results found for **\"{keyword}\"** in the uploaded documents.")
else:
    st.info("👆 Upload one or more Word (.docx) files to start searching.")
