"""
Word Document Search Tool
Drop Word, PDF, or Excel files and search by keyword to find all related content.
"""
import html
import re
import streamlit as st
from search_engine import (
    EXCEL_ROW_SEP,
    extract_text_from_docx,
    extract_text_from_pdf,
    extract_text_from_excel,
    extract_images_from_docx,
    get_excel_headers,
    get_nearest_image,
    get_word_suggestions,
    search_keyword,
)

st.set_page_config(page_title="Word Search Tool", page_icon="📄", layout="wide", initial_sidebar_state="expanded")

# Custom CSS for a cleaner, more attractive look
st.markdown("""
<style>
    /* Main container and typography */
    .stApp { background: linear-gradient(180deg, #f8fafc 0%, #f1f5f9 100%); }
    h1 { color: #1e293b !important; font-weight: 700 !important; letter-spacing: -0.02em !important; }
    
    /* Success message - pill style */
    .stSuccess { border-radius: 12px !important; background: linear-gradient(135deg, #e0f2fe 0%, #bae6fd 100%) !important; border: none !important; }
    
    /* Metric cards */
    [data-testid="stMetricValue"] { font-size: 1.75rem !important; font-weight: 700 !important; color: #0f172a !important; }
    [data-testid="stMetricLabel"] { color: #64748b !important; font-weight: 500 !important; }
    
    /* Expander header */
    .streamlit-expanderHeader { background: linear-gradient(90deg, #f1f5f9 0%, #e2e8f0 100%) !important; border-radius: 10px !important; padding: 0.75rem 1rem !important; }
    
    /* Suggestion area */
    div[data-testid="stVerticalBlock"] > div:has(button) { margin-bottom: 0.5rem; }
    .stButton > button { border-radius: 20px !important; font-weight: 500 !important; transition: all 0.2s !important; }
    .stButton > button:hover { transform: translateY(-1px); box-shadow: 0 4px 12px rgba(59, 130, 246, 0.3) !important; }
    
    /* Info message */
    .stInfo { border-radius: 12px !important; border-left: 4px solid #3b82f6 !important; }
</style>
""", unsafe_allow_html=True)

st.title("📄 Word Document Search")
st.caption("Upload Word (.docx), PDF, or Excel (.xlsx, .xls) files, then search by keyword.")

# File uploader - Word, PDF, Excel
uploaded_files = st.file_uploader(
    "Drop your files here",
    type=["docx", "pdf", "xlsx", "xls"],
    accept_multiple_files=True,
    help="Select .docx, .pdf, .xlsx, or .xls files",
)

# Store extracted content in session so we don't re-parse on every keystroke
if "all_paragraphs" not in st.session_state:
    st.session_state.all_paragraphs = []
if "all_images" not in st.session_state:
    st.session_state.all_images = []
if "file_bytes_by_name" not in st.session_state:
    st.session_state.file_bytes_by_name = {}
if "excel_headers" not in st.session_state:
    st.session_state.excel_headers = {}  # (filename, sheet_name) -> [col1, col2, ...]

if uploaded_files:
    file_ids = [f.name + str(f.size) for f in uploaded_files]
    if "last_file_ids" not in st.session_state or st.session_state.last_file_ids != file_ids:
        st.session_state.last_file_ids = file_ids
        all_paragraphs = []
        all_images = []
        file_bytes_by_name = {}
        excel_headers = {}
        for f in uploaded_files:
            b = f.getvalue()
            file_bytes_by_name[f.name] = b
            name_lower = f.name.lower()
            if name_lower.endswith(".docx"):
                all_paragraphs.extend(extract_text_from_docx(b, f.name))
                all_images.extend(extract_images_from_docx(b, f.name))
            elif name_lower.endswith(".pdf"):
                all_paragraphs.extend(extract_text_from_pdf(b, f.name))
            elif name_lower.endswith((".xlsx", ".xls")):
                all_paragraphs.extend(extract_text_from_excel(b, f.name))
                excel_headers.update(get_excel_headers(b, f.name))
        st.session_state.all_paragraphs = all_paragraphs
        st.session_state.all_images = all_images
        st.session_state.file_bytes_by_name = file_bytes_by_name
        st.session_state.excel_headers = excel_headers

    total_images = len(st.session_state.all_images)
    st.success(f"Loaded **{len(uploaded_files)}** file(s) • **{len(st.session_state.all_paragraphs)}** text blocks • **{total_images}** image(s) indexed.")

    # Apply clicked suggestion before the keyword widget is created (Streamlit rule)
    if "suggestion_clicked" in st.session_state:
        st.session_state["keyword"] = st.session_state.pop("suggestion_clicked")

    # Search box
    keyword = st.text_input(
        "Search keyword",
        placeholder="Type any word or phrase...",
        key="keyword",
    )

    # Word suggestions from documents (click to use as keyword)
    all_paragraphs = st.session_state.all_paragraphs
    if all_paragraphs:
        prefix = (keyword or "").strip()
        suggestions = get_word_suggestions(all_paragraphs, prefix=prefix, max_suggestions=15)
        if suggestions:
            st.markdown("**✨ Word suggestions** — *click to search*")
            cols = st.columns(5)
            for i, word in enumerate(suggestions[:15]):
                with cols[i % 5]:
                    if st.button(word, key=f"sugg_{word}_{i}"):
                        st.session_state["suggestion_clicked"] = word
                        st.rerun()
            st.markdown("---")

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
                # Highlighter-style: yellow background + bold so the search keyword stands out
                repl = lambda m: f'<span style="background-color:#fef08a;font-weight:700;padding:0 3px;border-radius:3px;">{m.group(0)}</span>'
                return re.sub(re.escape(keyword), repl, s, flags=re.IGNORECASE)

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
                # One exact line per occurrence (matching paragraphs only, in order)
                occurrences_exact = matching_texts_for_doc
                doc_imgs = images_by_doc.get(doc_name, [])
                nearest = None
                # Only Word (.docx) has image extraction and nearest-image logic
                if doc_name.lower().endswith(".docx"):
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
                    "occurrences_exact": occurrences_exact,
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
            st.markdown('<div style="margin: 1rem 0;">', unsafe_allow_html=True)
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("📁 Documents found", len(doc_results))
            with col2:
                st.metric("🔢 Total count", total_count)
            with col3:
                st.metric("🔍 Keyword", f"\"{keyword}\"")
            st.markdown("---")

            excel_headers = st.session_state.get("excel_headers", {})

            for r in doc_results:
                with st.expander(f"📄 **{r['name']}** — **{r['count']}** match(es)", expanded=True):
                    occs = r.get("occurrences_exact") or []
                    if occs:
                        st.markdown("**📌 Occurrences** — *scroll to read each exact line*")
                        for idx, exact_line in enumerate(occs, 1):
                            badge = f'<span style="background: linear-gradient(135deg, #3b82f6, #2563eb); color: white; padding: 4px 10px; border-radius: 12px; font-size: 0.85rem; font-weight: 600;">Occurrence {idx}</span>'
                            if EXCEL_ROW_SEP in exact_line:
                                parts = exact_line.split(EXCEL_ROW_SEP)
                                sheet_part = parts[0] if parts else ""
                                cells = parts[1:] if len(parts) > 1 else []
                                sheet_name = sheet_part.replace("[Sheet:", "").replace("]", "").strip()
                                headers = excel_headers.get((r["name"], sheet_name), [])
                                if len(headers) < len(cells):
                                    headers = headers + [f"Column {i+1}" for i in range(len(headers), len(cells))]
                                elif len(headers) > len(cells):
                                    headers = headers[:len(cells)]
                                cells_escaped = [html.escape(str(c)) for c in cells]
                                cells_highlighted = [highlight(c).replace("\n", "<br>") for c in cells_escaped]
                                th_html = "".join(
                                    f'<th style="padding:10px 14px;border:1px solid #cbd5e1;background:#f1f5f9;color:#0f172a;font-weight:600;text-align:left;">{html.escape(h)}</th>'
                                    for h in headers
                                )
                                td_html = "".join(
                                    f'<td style="padding:10px 14px;border:1px solid #e2e8f0;background:#fff;color:#334155;">{c}</td>'
                                    for c in cells_highlighted
                                )
                                content = f'<p style="color:#64748b;font-size:0.9rem;margin-bottom:8px;">{html.escape(sheet_part)}</p><table style="border-collapse:collapse;width:100%;"><thead><tr>{th_html}</tr></thead><tbody><tr>{td_html}</tr></tbody></table>'
                            else:
                                exact_highlighted = highlight(html.escape(exact_line)).replace("\n", "<br>")
                                content = f'<div style="color:#334155;line-height:1.6;">{exact_highlighted}</div>'
                            st.markdown(
                                f'<div style="background: linear-gradient(135deg, #ffffff 0%, #f8fafc 100%); padding: 14px 16px; border-radius: 10px; border-left: 4px solid #3b82f6; margin-bottom: 12px; box-shadow: 0 1px 3px rgba(0,0,0,0.06);">{badge}<div style="margin-top: 10px;">{content}</div></div>',
                                unsafe_allow_html=True,
                            )
                        st.markdown("---")
                    col_text, col_img = st.columns([2, 1])
                    with col_text:
                        st.markdown("**📋 Full context**")
                        if r["name"].lower().endswith((".xlsx", ".xls")) and EXCEL_ROW_SEP in r.get("complete_text", ""):
                            rows_html = []
                            header_row_done = {}
                            for row_text in r["complete_text"].split("\n\n"):
                                if EXCEL_ROW_SEP not in row_text:
                                    continue
                                parts = row_text.split(EXCEL_ROW_SEP)
                                sheet_part = parts[0] if parts else ""
                                sheet_name = sheet_part.replace("[Sheet:", "").replace("]", "").strip()
                                cells = parts[1:] if len(parts) > 1 else []
                                headers = excel_headers.get((r["name"], sheet_name), [])
                                if len(headers) < len(cells):
                                    headers = headers + [f"Column {i+1}" for i in range(len(headers), len(cells))]
                                elif len(headers) > len(cells):
                                    headers = headers[:len(cells)]
                                cells_highlighted = [highlight(html.escape(str(c))).replace("\n", "<br>") for c in cells]
                                if sheet_name not in header_row_done:
                                    th_html = "".join(
                                        f'<th style="padding:8px 12px;border:1px solid #cbd5e1;background:#f1f5f9;font-weight:600;">{html.escape(h)}</th>' for h in headers
                                    )
                                    rows_html.append(f'<tr>{th_html}</tr>')
                                    header_row_done[sheet_name] = True
                                td_html = "".join(
                                    f'<td style="padding:8px 12px;border:1px solid #e2e8f0;background:#fff;">{c}</td>' for c in cells_highlighted
                                )
                                rows_html.append(f'<tr>{td_html}</tr>')
                            full_table = f'<table style="border-collapse:collapse;width:100%;"><tbody>{"".join(rows_html)}</tbody></table>'
                            st.markdown(
                                f'<div style="background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%); padding: 16px 18px; border-radius: 10px; border-left: 4px solid #3b82f6; box-shadow: 0 1px 3px rgba(0,0,0,0.06);">{full_table}</div>',
                                unsafe_allow_html=True,
                            )
                        else:
                            st.markdown(
                                f'<div style="background: linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%); padding: 16px 18px; border-radius: 10px; border-left: 4px solid #3b82f6; box-shadow: 0 1px 3px rgba(0,0,0,0.06); color: #334155; line-height: 1.65;">{r["highlighted"]}</div>',
                                unsafe_allow_html=True,
                            )
                    with col_img:
                        if show_images and r["nearest"] and r["name"].lower().endswith(".docx"):
                            st.markdown("**🖼️ Relevant image**")
                            img_bytes, img_name = r["nearest"]
                            try:
                                st.image(img_bytes, caption=img_name, use_container_width=True)
                            except Exception:
                                st.caption(f"*{img_name}*")
            st.markdown("---")
        else:
            st.markdown(
                f'<div style="background: linear-gradient(135deg, #fef3c7 0%, #fde68a 100%); padding: 1rem 1.25rem; border-radius: 12px; border-left: 4px solid #f59e0b;">'
                f'<strong>No results found</strong> for <em>"{html.escape(keyword)}"</em> in the uploaded documents. Try another keyword or check the word suggestions above.</div>',
                unsafe_allow_html=True,
            )
else:
    st.markdown(
        '<p style="color: #64748b; font-size: 1.05rem;">👆 Upload one or more <strong>Word</strong> (.docx), <strong>PDF</strong>, or <strong>Excel</strong> (.xlsx, .xls) files to start searching.</p>',
        unsafe_allow_html=True,
    )
