"""Extract text and images from Word docs and search by keyword."""
import xml.etree.ElementTree as ET
import zipfile
from io import BytesIO
from typing import List, Optional, Tuple
from docx import Document

# Word XML namespaces
NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}


def _get_text(elem) -> str:
    """Extract all text from an element (paragraph or cell)."""
    if elem is None:
        return ""
    # Handle any namespace on w:t
    texts = []
    for e in elem.iter():
        if e.tag.endswith("}t") or e.tag == "t":
            texts.append(e.text or "")
    return "".join(texts)


def _find_blip_embed(elem) -> Optional[str]:
    """Find r:embed (relationship id) for an image in drawing."""
    if elem is None:
        return None
    r_embed = "http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
    for e in elem.iter():
        if "blip" in (e.tag or "").lower():
            for k, v in (e.attrib or {}).items():
                if "embed" in k.lower() and v:
                    return v
    return None


def get_nearest_image(
    file_bytes: bytes,
    filename: str,
    matching_texts: List[str],
    images_list: List[Tuple[str, bytes, str]],
) -> Optional[Tuple[bytes, str]]:
    """
    From the document, find the image that appears closest to text matching the keyword.
    images_list: list of (doc_name, image_bytes, image_name) for this document.
    Returns (image_bytes, image_name) or None if no image / no match.
    """
    if not matching_texts or not images_list:
        return (images_list[0][1], images_list[0][2]) if images_list else None
    doc_images = [(b, n) for fn, b, n in images_list if fn == filename]
    if not doc_images:
        return None
    try:
        with zipfile.ZipFile(BytesIO(file_bytes), "r") as z:
            doc_xml = z.read("word/document.xml")
            try:
                rels_xml = z.read("word/_rels/document.xml.rels")
            except KeyError:
                return doc_images[0]
        root = ET.fromstring(doc_xml)
        body = next((c for c in root if c.tag and "body" in (c.tag or "").lower()), None)
        if body is None:
            return doc_images[0]
        # Build (position, type, value): 'text' -> text, 'image' -> rId
        blocks = []
        rid_to_target = {}
        try:
            rels_root = ET.fromstring(rels_xml)
            for rel in rels_root.findall(".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
                r_id = rel.get("Id")
                target = rel.get("Target", "")
                if "media/" in target:
                    rid_to_target[r_id] = "word/" + target.lstrip("/")
        except Exception:
            pass
        pos = 0
        for child in body:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag == "p":
                para_text = _get_text(child).strip()
                if para_text:
                    blocks.append((pos, "text", para_text))
                    pos += 1
                embed = _find_blip_embed(child)
                if embed:
                    blocks.append((pos, "image", embed))
                    pos += 1
            elif tag == "tbl":
                for row in child:
                    if row.tag.endswith("}tr") or row.tag == "tr":
                        for cell in row:
                            if cell.tag.endswith("}tc") or cell.tag == "tc":
                                cell_text = _get_text(cell).strip()
                                if cell_text:
                                    blocks.append((pos, "text", cell_text))
                                    pos += 1
                for row in child:
                    if row.tag.endswith("}tr") or row.tag == "tr":
                        for cell in row:
                            if cell.tag.endswith("}tc") or cell.tag == "tc":
                                embed = _find_blip_embed(cell)
                                if embed:
                                    blocks.append((pos, "image", embed))
                                    pos += 1
        image_blocks = [(i, v) for i, (_, t, v) in enumerate(blocks) if t == "image"]
        if not image_blocks:
            return doc_images[0]
        matching_indices = set()
        for i, (_, typ, block_text) in enumerate(blocks):
            if typ != "text" or not block_text:
                continue
            if any(mt.strip().lower() in block_text.lower() for mt in matching_texts if mt.strip()):
                matching_indices.add(i)
        if not matching_indices:
            return doc_images[0]
        def dist(img_idx):
            return min(abs(img_idx - ti) for ti in matching_indices)
        nearest_img_idx, nearest_rid = min(image_blocks, key=lambda x: dist(x[0]))
        target_path = rid_to_target.get(nearest_rid)
        if target_path:
            with zipfile.ZipFile(BytesIO(file_bytes), "r") as z:
                try:
                    img_bytes = z.read(target_path)
                    name = target_path.split("/")[-1]
                    return (img_bytes, name)
                except Exception:
                    pass
        return doc_images[0]
    except Exception:
        return doc_images[0] if doc_images else None


def extract_images_from_docx(file_bytes: bytes, filename: str) -> List[Tuple[str, bytes, str]]:
    """
    Extract all images from a .docx file (docx is a zip with word/media/).
    Returns list of (source_doc_name, image_bytes, image_filename).
    """
    results = []
    try:
        with zipfile.ZipFile(BytesIO(file_bytes), "r") as z:
            for name in z.namelist():
                if name.startswith("word/media/"):
                    image_name = name.split("/")[-1]
                    results.append((filename, z.read(name), image_name))
    except Exception:
        pass
    return results


def extract_text_from_docx(file_bytes: bytes, filename: str) -> List[Tuple[str, str]]:
    """
    Extract all paragraphs from a .docx file.
    Returns list of (filename, paragraph_text).
    """
    results = []
    try:
        doc = Document(BytesIO(file_bytes))
        for para in doc.paragraphs:
            text = para.text.strip()
            if text:
                results.append((filename, text))
        # Also extract text from tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    text = cell.text.strip()
                    if text:
                        results.append((filename, text))
    except Exception as e:
        results.append((filename, f"[Error reading file: {e}]"))
    return results


def search_keyword(paragraphs: List[Tuple[str, str]], keyword: str) -> List[Tuple[str, str]]:
    """
    Find all (filename, paragraph) where paragraph contains keyword (case-insensitive).
    """
    if not keyword or not keyword.strip():
        return []
    k = keyword.strip().lower()
    return [(fname, text) for fname, text in paragraphs if k in text.lower()]
