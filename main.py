from fastapi import FastAPI, UploadFile, File
from docx import Document
import tempfile

app = FastAPI()


def extract_text(doc):
    return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])


def extract_formatting(doc):
    section = doc.sections[0]

    def pt(x):
        return x.pt if x else None

    # margins (in points)
    margins = {
        "top_pt": pt(section.top_margin),
        "bottom_pt": pt(section.bottom_margin),
        "left_pt": pt(section.left_margin),
        "right_pt": pt(section.right_margin),
    }

    # font detection (best-effort)
    font_name = None
    font_size = None

    for p in doc.paragraphs:
        for r in p.runs:
            if r.font.name:
                font_name = r.font.name
            if r.font.size:
                font_size = r.font.size.pt
            if font_name and font_size:
                break

    # paragraph formatting (first paragraph sample)
    line_spacing = None
    indent = None

    if doc.paragraphs:
        p = doc.paragraphs[0]
        if p.paragraph_format.line_spacing:
            line_spacing = p.paragraph_format.line_spacing
        if p.paragraph_format.first_line_indent:
            indent = p.paragraph_format.first_line_indent.pt

    return {
        "margins": margins,
        "font": {
            "name": font_name,
            "size_pt": font_size
        },
        "line_spacing": line_spacing,
        "indentation_pt": indent
    }


@app.post("/convert-docx")
async def convert_docx(file: UploadFile = File(...)):

    if not file.filename.endswith(".docx"):
        return {"error": "Only DOCX allowed"}

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp.write(await file.read())
        path = tmp.name

    doc = Document(path)

    text = extract_text(doc)
    formatting = extract_formatting(doc)

    return {
        "filename": file.filename,
        "text": text,
        "word_count": len(text.split()),
        "formatting": formatting
    }