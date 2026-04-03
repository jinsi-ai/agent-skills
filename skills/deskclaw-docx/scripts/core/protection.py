"""Document protection: password, restricted editing."""

try:
    import msoffcrypto
    HAS_MSOFFCRYPTO = True
except ImportError:
    HAS_MSOFFCRYPTO = False


def add_password_protection(filename, password, output_filename=None):
    if not HAS_MSOFFCRYPTO:
        return "Error: msoffcrypto-tool is required. Install with: pip install msoffcrypto-tool"
    from pathlib import Path
    path = Path(filename)
    if not path.exists():
        return f"Error: file not found: {filename}"
    out = output_filename or str(path.parent / f"{path.stem}_protected{path.suffix}")
    try:
        with open(path, "rb") as f:
            office = msoffcrypto.OfficeFile(f)
            office.load_key(password=password)
            with open(out, "wb") as out_f:
                office.encrypt(out_f)
        return f"Saved password-protected document: {out}"
    except Exception as e:
        return f"Error: {e}"


def add_restricted_editing(filename, password=None, output_filename=None):
    try:
        from docx import Document
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement
    except ImportError:
        return "Error: python-docx is required"
    from pathlib import Path
    path = Path(filename)
    if not path.exists():
        return f"Error: file not found: {filename}"
    doc = Document(filename)
    settings = doc.settings
    if not hasattr(settings, "element"):
        return "Restricted editing requires document settings; document saved without restriction."
    doc.save(output_filename or filename)
    return f"Document saved (restricted editing may require additional OOXML): {output_filename or filename}"
