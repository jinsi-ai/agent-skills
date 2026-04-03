"""Comment extraction from Word documents."""

import json
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def _get_comment_ids_to_text(docx_path):
    """Parse word/comments.xml and return comment_id -> {author, date, text}."""
    result = {}
    with zipfile.ZipFile(docx_path, "r") as z:
        if "word/comments.xml" not in z.namelist():
            return result
        with z.open("word/comments.xml") as f:
            tree = ET.parse(f)
            root = tree.getroot()
            for c in root.findall(".//w:comment", NS):
                cid = c.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id")
                author = c.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author", "")
                date = c.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date", "")
                texts = []
                for t in c.findall(".//w:t", NS):
                    if t.text:
                        texts.append(t.text)
                result[cid] = {"author": author, "date": date, "text": "".join(texts)}
    return result


def get_all_comments(filename):
    path = Path(filename)
    if not path.exists():
        return json.dumps({"error": f"File not found: {filename}"})
    comments = _get_comment_ids_to_text(path)
    out = [{"id": k, **v} for k, v in sorted(comments.items(), key=lambda x: int(x[0]))]
    return json.dumps(out, ensure_ascii=False, indent=2)


def get_comments_by_author(filename, author):
    path = Path(filename)
    if not path.exists():
        return json.dumps({"error": f"File not found: {filename}"})
    comments = _get_comment_ids_to_text(path)
    out = [{"id": k, **v} for k, v in comments.items() if v["author"] == author]
    return json.dumps(out, ensure_ascii=False, indent=2)


def get_comments_for_paragraph(filename, paragraph_index):
    path = Path(filename)
    if not path.exists():
        return json.dumps({"error": f"File not found: {filename}"})
    comments = _get_comment_ids_to_text(path)
    try:
        from docx import Document
        doc = Document(filename)
        paras = list(doc.paragraphs)
        if paragraph_index < 0 or paragraph_index >= len(paras):
            return json.dumps({"error": "paragraph_index out of range"})
        p = paras[paragraph_index]
        p_xml = p._p
        comment_refs = p_xml.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}commentReference")
        ids = [ref.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id") for ref in comment_refs]
        out = [{"id": i, **comments.get(i, {"author": "", "date": "", "text": ""})} for i in ids if i in comments]
        return json.dumps(out, ensure_ascii=False, indent=2)
    except ImportError:
        return get_all_comments(filename)


def add_comment(filename, paragraph_index, comment_text, author="AI Assistant", start_pos=None, end_pos=None):
    """Add a comment to a paragraph.
    
    paragraph_index: which paragraph to add comment to
    comment_text: the comment content
    author: comment author name
    start_pos: start position in paragraph text (optional, comments whole paragraph if not specified)
    end_pos: end position in paragraph text (optional)
    
    Note: This creates a new comment by modifying the document's XML structure.
    """
    from datetime import datetime
    import shutil
    import tempfile
    import os
    
    path = Path(filename)
    if not path.exists():
        return f"Error: File not found: {filename}"
    
    # Get next comment ID
    existing_comments = _get_comment_ids_to_text(path)
    next_id = str(max([int(k) for k in existing_comments.keys()] + [0]) + 1)
    
    # Current datetime in ISO format
    now = datetime.now().strftime("%Y-%m-%dT%H:%M:%SZ")
    
    # Create temporary directory
    with tempfile.TemporaryDirectory() as tmpdir:
        # Extract docx
        with zipfile.ZipFile(path, 'r') as z:
            z.extractall(tmpdir)
        
        # Check if comments.xml exists, create if not
        comments_path = os.path.join(tmpdir, "word", "comments.xml")
        if not os.path.exists(comments_path):
            # Create new comments.xml
            comments_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
</w:comments>'''
            with open(comments_path, 'w', encoding='utf-8') as f:
                f.write(comments_xml)
            
            # Update [Content_Types].xml to include comments
            content_types_path = os.path.join(tmpdir, "[Content_Types].xml")
            tree = ET.parse(content_types_path)
            root = tree.getroot()
            
            # Check if comments override already exists
            ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
            has_comments = False
            for override in root.findall(f".//{{{ct_ns}}}Override"):
                if override.get("PartName") == "/word/comments.xml":
                    has_comments = True
                    break
            
            if not has_comments:
                override = ET.SubElement(root, f"{{{ct_ns}}}Override")
                override.set("PartName", "/word/comments.xml")
                override.set("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml")
                tree.write(content_types_path, xml_declaration=True, encoding='UTF-8')
            
            # Update document.xml.rels
            rels_path = os.path.join(tmpdir, "word", "_rels", "document.xml.rels")
            rels_tree = ET.parse(rels_path)
            rels_root = rels_tree.getroot()
            
            # Find next rId
            rels_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
            max_rid = 0
            for rel in rels_root.findall(f".//{{{rels_ns}}}Relationship"):
                rid = rel.get("Id", "rId0")
                try:
                    rid_num = int(rid.replace("rId", ""))
                    max_rid = max(max_rid, rid_num)
                except ValueError:
                    pass
            
            new_rid = f"rId{max_rid + 1}"
            rel = ET.SubElement(rels_root, f"{{{rels_ns}}}Relationship")
            rel.set("Id", new_rid)
            rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments")
            rel.set("Target", "comments.xml")
            rels_tree.write(rels_path, xml_declaration=True, encoding='UTF-8')
        
        # Add comment to comments.xml
        tree = ET.parse(comments_path)
        root = tree.getroot()
        
        # Register namespace
        ET.register_namespace('w', NS['w'])
        
        # Create comment element
        comment = ET.SubElement(root, f"{{{NS['w']}}}comment")
        comment.set(f"{{{NS['w']}}}id", next_id)
        comment.set(f"{{{NS['w']}}}author", author)
        comment.set(f"{{{NS['w']}}}date", now)
        comment.set(f"{{{NS['w']}}}initials", author[:2].upper() if author else "AI")
        
        # Add paragraph with text
        p = ET.SubElement(comment, f"{{{NS['w']}}}p")
        r = ET.SubElement(p, f"{{{NS['w']}}}r")
        t = ET.SubElement(r, f"{{{NS['w']}}}t")
        t.text = comment_text
        
        tree.write(comments_path, xml_declaration=True, encoding='UTF-8')
        
        # Add comment reference to document.xml
        doc_path = os.path.join(tmpdir, "word", "document.xml")
        doc_tree = ET.parse(doc_path)
        doc_root = doc_tree.getroot()
        
        # Find the paragraph
        paragraphs = doc_root.findall(f".//{{{NS['w']}}}p")
        if paragraph_index >= len(paragraphs):
            return f"Error: paragraph_index {paragraph_index} out of range"
        
        target_p = paragraphs[paragraph_index]
        
        # Add comment range start at the beginning of paragraph
        comment_start = ET.Element(f"{{{NS['w']}}}commentRangeStart")
        comment_start.set(f"{{{NS['w']}}}id", next_id)
        target_p.insert(0, comment_start)
        
        # Add comment range end and reference at the end
        comment_end = ET.Element(f"{{{NS['w']}}}commentRangeEnd")
        comment_end.set(f"{{{NS['w']}}}id", next_id)
        target_p.append(comment_end)
        
        # Add run with comment reference
        r = ET.Element(f"{{{NS['w']}}}r")
        comment_ref = ET.SubElement(r, f"{{{NS['w']}}}commentReference")
        comment_ref.set(f"{{{NS['w']}}}id", next_id)
        target_p.append(r)
        
        doc_tree.write(doc_path, xml_declaration=True, encoding='UTF-8')
        
        # Repack the docx
        output_path = str(path) + ".tmp"
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as z:
            for root_dir, dirs, files in os.walk(tmpdir):
                for file in files:
                    file_path = os.path.join(root_dir, file)
                    arcname = os.path.relpath(file_path, tmpdir)
                    z.write(file_path, arcname)
        
        # Replace original file
        shutil.move(output_path, path)
    
    return f"Added comment to paragraph {paragraph_index}: {comment_text[:30]}..."


def delete_comment(filename, comment_id):
    """Delete a comment by its ID."""
    import shutil
    import tempfile
    import os
    
    path = Path(filename)
    if not path.exists():
        return f"Error: File not found: {filename}"
    
    comment_id = str(comment_id)
    
    with tempfile.TemporaryDirectory() as tmpdir:
        # Extract docx
        with zipfile.ZipFile(path, 'r') as z:
            z.extractall(tmpdir)
        
        # Remove from comments.xml
        comments_path = os.path.join(tmpdir, "word", "comments.xml")
        if os.path.exists(comments_path):
            tree = ET.parse(comments_path)
            root = tree.getroot()
            
            for comment in root.findall(f".//{{{NS['w']}}}comment"):
                if comment.get(f"{{{NS['w']}}}id") == comment_id:
                    root.remove(comment)
                    break
            
            tree.write(comments_path, xml_declaration=True, encoding='UTF-8')
        
        # Remove references from document.xml
        doc_path = os.path.join(tmpdir, "word", "document.xml")
        doc_tree = ET.parse(doc_path)
        doc_root = doc_tree.getroot()
        
        # Remove commentRangeStart, commentRangeEnd, and commentReference
        for elem in doc_root.iter():
            to_remove = []
            for child in elem:
                tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                if tag in ('commentRangeStart', 'commentRangeEnd'):
                    if child.get(f"{{{NS['w']}}}id") == comment_id:
                        to_remove.append(child)
                elif tag == 'r':
                    for ref in child.findall(f".//{{{NS['w']}}}commentReference"):
                        if ref.get(f"{{{NS['w']}}}id") == comment_id:
                            to_remove.append(child)
                            break
            for item in to_remove:
                elem.remove(item)
        
        doc_tree.write(doc_path, xml_declaration=True, encoding='UTF-8')
        
        # Repack
        output_path = str(path) + ".tmp"
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as z:
            for root_dir, dirs, files in os.walk(tmpdir):
                for file in files:
                    file_path = os.path.join(root_dir, file)
                    arcname = os.path.relpath(file_path, tmpdir)
                    z.write(file_path, arcname)
        
        shutil.move(output_path, path)
    
    return f"Deleted comment {comment_id}"
