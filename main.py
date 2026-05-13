from fastapi import FastAPI, UploadFile, File
from docx import Document
import tempfile

app = FastAPI()


def extract_text(doc):
    return "\n".join([p.text for p in doc.paragraphs if p.text and p.text.strip()])


def extract_font(doc):
    font_name = None
    font_size = None

    for p in doc.paragraphs:
        for r in p.runs:
            if r.font and r.font.name:
                font_name = r.font.name
            if r.font and r.font.size:
                font_size = r.font.size.pt

            if font_name and font_size:
                return font_name, font_size

    return font_name, font_size


def extract_line_spacing(paragraphs):
    values = []

    for p in paragraphs[:10]:  # sample more than 1 paragraph
        fmt = p.paragraph_format
        if fmt.line_spacing is not None:
            try:
                values.append(float(fmt.line_spacing))
            except:
                pass

    return values[0] if values else None


def extract_formatting(doc):
    section = doc.sections[0]

    def pt(x):
        return x.pt if x else None

    # margins
    margins = {
        "top_pt": pt(section.top_margin),
        "bottom_pt": pt(section.bottom_margin),
        "left_pt": pt(section.left_margin),
        "right_pt": pt(section.right_margin),
    }

    # font (FIXED)
    font_name, font_size = extract_font(doc)

    # line spacing (FIXED)
    line_spacing = extract_line_spacing(doc.paragraphs)

    # indentation (still sample-based but safer)
    indentation = {
        "first_line_pt": None,
        "left_pt": None,
        "right_pt": None
    }

    for p in doc.paragraphs[:5]:
        fmt = p.paragraph_format

        if fmt.first_line_indent:
            indentation["first_line_pt"] = fmt.first_line_indent.pt

        if fmt.left_indent:
            indentation["left_pt"] = fmt.left_indent.pt

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
        return {"success": False, "error": "Only DOCX allowed"}

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp.write(await file.read())
        path = tmp.name

    doc = Document(path)

    text = extract_text(doc)

    formatting = extract_formatting(doc)

    word_count = len(text.split())

    # 🔴 IMPORTANT: detect if file is actually readable
    formatting_complete = all([
        formatting["font"]["name"] is not None,
        formatting["font"]["size_pt"] is not None,
        word_count > 0
    ])

    return {
        "success": True,
        "filename": file.filename,
        "text": text,
        "word_count": word_count,
        "formatting": formatting,
        "formatting_complete": formatting_complete
    }
