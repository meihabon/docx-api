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

    indentation = {
        "first_line_pt": None,
        "left_pt": None,
        "right_pt": None
    }

    if doc.paragraphs:
        p = doc.paragraphs[0]
        fmt = p.paragraph_format

        # line spacing
        if fmt.line_spacing:
            try:
                line_spacing = float(fmt.line_spacing)
            except:
                line_spacing = str(fmt.line_spacing)

        # first line indent
        if fmt.first_line_indent:
            indentation["first_line_pt"] = fmt.first_line_indent.pt

        # left indent
        if fmt.left_indent:
            indentation["left_pt"] = fmt.left_indent.pt

        # right indent
        if fmt.right_indent:
            indentation["right_pt"] = fmt.right_indent.pt

    return {
        "margins": margins,
        "font": {
            "name": font_name,
            "size_pt": font_size
        },
        "line_spacing": line_spacing,
        "indentation": indentation
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
