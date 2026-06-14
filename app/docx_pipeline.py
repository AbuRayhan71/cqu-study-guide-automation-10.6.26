from __future__ import annotations

import copy
import re
import zipfile
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path
from typing import Iterable, Literal

from docx import Document
from docx.document import Document as DocumentObject
from docx.oxml.ns import qn
from docx.shared import Inches
from docx.table import Table
from docx.text.paragraph import Paragraph


BlockType = Literal["heading", "paragraph", "bullet_list", "numbered_list", "table", "image", "quote"]


@dataclass
class GenerationMetadata:
    unit_code: str = ""
    unit_name: str = ""
    week_number: str = ""
    week_title: str = ""
    version: str = ""
    year: str = ""
    term: str = ""


@dataclass
class ContentBlock:
    type: BlockType
    text: str = ""
    level: int = 1
    number: str = ""
    items: list[str] = field(default_factory=list)
    rows: list[list[str]] = field(default_factory=list)
    image_id: str = ""
    caption: str = ""


@dataclass
class ExtractedImage:
    image_id: str
    blob: bytes
    filename: str
    content_type: str


@dataclass
class StructuredDocument:
    metadata: GenerationMetadata
    body_blocks: list[ContentBlock]
    images: dict[str, ExtractedImage]


@dataclass
class GenerationResult:
    output_path: Path
    block_count: int
    image_count: int
    table_count: int


NAMESPACES = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def iter_block_items(parent: DocumentObject) -> Iterable[Paragraph | Table]:
    body = parent.element.body
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, parent)
        elif child.tag == qn("w:tbl"):
            yield Table(child, parent)


def extract_docx(path: Path, overrides: GenerationMetadata | None = None) -> StructuredDocument:
    document = Document(path)
    blocks: list[ContentBlock] = []
    images: dict[str, ExtractedImage] = {}
    image_counter = 1

    for item in iter_block_items(document):
        if isinstance(item, Paragraph):
            text = clean_text(item.text)
            style_name = item.style.name if item.style else ""
            for rid in image_relationship_ids(item):
                image_id = f"img_{image_counter:03d}"
                image_counter += 1
                related = document.part.related_parts[rid]
                filename = Path(getattr(related, "partname", f"{image_id}.bin")).name
                images[image_id] = ExtractedImage(
                    image_id=image_id,
                    blob=related.blob,
                    filename=filename,
                    content_type=getattr(related, "content_type", "application/octet-stream"),
                )
                blocks.append(ContentBlock(type="image", image_id=image_id))
            if not text:
                continue
            if style_name.lower().startswith("heading"):
                blocks.append(ContentBlock(type="heading", level=heading_level(style_name), text=text))
            elif is_numbered_paragraph(item):
                blocks.append(ContentBlock(type="numbered_list", items=[text]))
            elif is_bullet_paragraph(item):
                blocks.append(ContentBlock(type="bullet_list", items=[text]))
            else:
                blocks.append(ContentBlock(type="paragraph", text=text))
        else:
            rows = [[clean_text(cell.text) for cell in row.cells] for row in item.rows]
            if rows:
                blocks.append(ContentBlock(type="table", rows=rows))

    metadata = infer_metadata(blocks)
    if overrides:
        metadata = merge_metadata(metadata, overrides)
    return StructuredDocument(metadata=metadata, body_blocks=coalesce_lists(blocks), images=images)


def generate_study_guide(
    input_docx: Path,
    template_docx: Path,
    output_docx: Path,
    metadata: GenerationMetadata | None = None,
) -> GenerationResult:
    structured = extract_docx(input_docx, metadata)
    render_document(structured, template_docx, output_docx)
    validate_output(output_docx, structured)
    return GenerationResult(
        output_path=output_docx,
        block_count=len(structured.body_blocks),
        image_count=len(structured.images),
        table_count=sum(1 for block in structured.body_blocks if block.type == "table"),
    )


def clean_text(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def image_relationship_ids(paragraph: Paragraph) -> list[str]:
    ids: list[str] = []
    for blip in paragraph._element.xpath(".//a:blip"):
        rid = blip.get(qn("r:embed"))
        if rid:
            ids.append(rid)
    return ids


def heading_level(style_name: str) -> int:
    match = re.search(r"(\d+)", style_name)
    if not match:
        return 1
    return max(1, min(9, int(match.group(1))))


def is_bullet_paragraph(paragraph: Paragraph) -> bool:
    style_name = paragraph.style.name.lower() if paragraph.style else ""
    return "bullet" in style_name


def is_numbered_paragraph(paragraph: Paragraph) -> bool:
    style_name = paragraph.style.name.lower() if paragraph.style else ""
    return "number" in style_name or "list paragraph" in style_name


def coalesce_lists(blocks: list[ContentBlock]) -> list[ContentBlock]:
    coalesced: list[ContentBlock] = []
    for block in blocks:
        if block.type in {"bullet_list", "numbered_list"} and coalesced and coalesced[-1].type == block.type:
            coalesced[-1].items.extend(block.items)
        else:
            coalesced.append(block)
    return coalesced


def infer_metadata(blocks: list[ContentBlock]) -> GenerationMetadata:
    joined = "\n".join(block.text for block in blocks if block.text)
    unit_code = first_match(joined, r"\b[A-Z]{4}\d{5}\b")
    week_number = first_match(joined, r"\bWEEK\s+(\d{1,2})\b", group=1)
    unit_name = ""
    for block in blocks[:10]:
        if block.text and block.text != unit_code and not re.search(r"\bSTUDY GUIDE\b", block.text, re.I):
            unit_name = block.text
            break
    return GenerationMetadata(unit_code=unit_code, unit_name=unit_name, week_number=week_number)


def first_match(text: str, pattern: str, group: int = 0) -> str:
    match = re.search(pattern, text, re.I)
    return match.group(group).strip() if match else ""


def merge_metadata(base: GenerationMetadata, override: GenerationMetadata) -> GenerationMetadata:
    return GenerationMetadata(
        unit_code=override.unit_code or base.unit_code,
        unit_name=override.unit_name or base.unit_name,
        week_number=override.week_number or base.week_number,
        week_title=override.week_title or base.week_title,
        version=override.version or base.version,
        year=override.year or base.year,
        term=override.term or base.term,
    )


def render_document(structured: StructuredDocument, template_path: Path, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    document = Document(template_path)
    replace_placeholders(document, structured.metadata)
    body_placeholder = find_paragraph_containing(document, "{{ body_subdoc }}")
    if body_placeholder is None:
        body_placeholder = document.add_paragraph()
    clear_paragraph(body_placeholder)

    anchor = body_placeholder
    for block in structured.body_blocks:
        anchor = insert_block_after(anchor, block, document, structured.images)

    save_and_reload(document, output_path)


def replace_placeholders(document: DocumentObject, metadata: GenerationMetadata) -> None:
    replacements = {
        "{{ unit_code }}": metadata.unit_code or "UNIT CODE",
        "{{ unit_name | upper }}": (metadata.unit_name or "UNIT NAME").upper(),
        "STUDY GUIDE | WEEK {{ week_number }}": f"STUDY GUIDE | WEEK {metadata.week_number or '1'}",
        "{{ week_title }}": metadata.week_title,
        "{{ version }}": metadata.version,
        "{{ toc_placeholder }}": "Update this table of contents in Word after opening the generated document.",
    }
    for paragraph in document.paragraphs:
        text = paragraph.text
        normalized = text.replace("{{ toc _placeholder }}", "{{ toc_placeholder }}")
        normalized = normalized.replace("{{ body _subdoc }}", "{{ body_subdoc }}")
        if normalized != text:
            set_paragraph_text(paragraph, normalized)
            text = normalized
        for placeholder, replacement in replacements.items():
            if placeholder in text:
                set_paragraph_text(paragraph, text.replace(placeholder, replacement))
                text = paragraph.text


def find_paragraph_containing(document: DocumentObject, needle: str) -> Paragraph | None:
    for paragraph in document.paragraphs:
        normalized = paragraph.text.replace("{{ body _subdoc }}", "{{ body_subdoc }}")
        if needle in normalized:
            return paragraph
    return None


def clear_paragraph(paragraph: Paragraph) -> None:
    for child in list(paragraph._p):
        paragraph._p.remove(child)


def set_paragraph_text(paragraph: Paragraph, text: str) -> None:
    clear_paragraph(paragraph)
    paragraph.add_run(text)


def insert_block_after(
    anchor: Paragraph,
    block: ContentBlock,
    document: DocumentObject,
    images: dict[str, ExtractedImage],
) -> Paragraph:
    if block.type == "heading":
        paragraph = insert_paragraph_after(anchor, block.text, style=f"Heading {block.level}")
        return paragraph
    if block.type == "paragraph":
        return insert_paragraph_after(anchor, block.text)
    if block.type == "quote":
        paragraph = insert_paragraph_after(anchor, block.text)
        paragraph.style = "Quote" if "Quote" in [style.name for style in document.styles] else paragraph.style
        if block.citation:
            return insert_paragraph_after(paragraph, block.citation)
        return paragraph
    if block.type in {"bullet_list", "numbered_list"}:
        current = anchor
        style = "List Bullet" if block.type == "bullet_list" else "List Number"
        for item in block.items:
            current = insert_paragraph_after(current, item, style=style)
        return current
    if block.type == "table":
        table = document.add_table(rows=0, cols=max(len(row) for row in block.rows))
        table.style = "Table Grid"
        for row_values in block.rows:
            cells = table.add_row().cells
            for idx, value in enumerate(row_values):
                cells[idx].text = value
        anchor._p.addnext(table._tbl)
        marker = insert_paragraph_after(anchor, "")
        table._tbl.addnext(marker._p)
        return marker
    if block.type == "image":
        paragraph = insert_paragraph_after(anchor, "")
        image = images.get(block.image_id)
        if image is None:
            raise ValueError(f"Missing extracted image {block.image_id}")
        run = paragraph.add_run()
        run.add_picture(BytesIO(image.blob), width=Inches(5.8))
        if block.caption:
            return insert_paragraph_after(paragraph, block.caption)
        return paragraph
    return anchor


def insert_paragraph_after(paragraph: Paragraph, text: str = "", style: str | None = None) -> Paragraph:
    new_p = copy.deepcopy(paragraph._p)
    for child in list(new_p):
        new_p.remove(child)
    paragraph._p.addnext(new_p)
    new_paragraph = Paragraph(new_p, paragraph._parent)
    if style:
        try:
            new_paragraph.style = style
        except KeyError:
            pass
    if text:
        new_paragraph.add_run(text)
    return new_paragraph


def save_and_reload(document: DocumentObject, output_path: Path) -> None:
    document.save(output_path)
    Document(output_path)


def validate_output(path: Path, structured: StructuredDocument) -> None:
    if not path.exists() or path.stat().st_size == 0:
        raise ValueError("Output DOCX was not created.")
    document = Document(path)
    text = "\n".join(paragraph.text for paragraph in document.paragraphs)
    unresolved = re.findall(r"{{[^}]+}}|{%[^%]+%}", text)
    if unresolved:
        raise ValueError(f"Unresolved template placeholders remain: {', '.join(unresolved)}")
    if any(block.type == "heading" for block in structured.body_blocks) and not any(
        paragraph.style and paragraph.style.name.lower().startswith("heading") for paragraph in document.paragraphs
    ):
        raise ValueError("Expected headings were not present in output.")
    expected_tables = sum(1 for block in structured.body_blocks if block.type == "table")
    if len(document.tables) < expected_tables:
        raise ValueError("One or more source tables were not present in output.")
    expected_images = len(structured.images)
    actual_images = count_docx_media(path)
    template_images = 4
    if expected_images and actual_images < template_images + expected_images:
        raise ValueError("One or more source images were not present in output.")


def count_docx_media(path: Path) -> int:
    with zipfile.ZipFile(path) as archive:
        return len([name for name in archive.namelist() if name.startswith("word/media/")])
