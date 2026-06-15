from __future__ import annotations

import json
import urllib.error
import urllib.request
from typing import Any

from app.config import settings
from app.docx_pipeline import ContentBlock, StructuredDocument


class AiNotConfiguredError(RuntimeError):
    pass


class AiResponseError(RuntimeError):
    pass


def analyze_document(document: StructuredDocument, hyperlinks: list[dict[str, str]]) -> dict[str, Any]:
    fallback = deterministic_analysis(document, hyperlinks)
    if not settings.enable_ai_polish:
        return fallback
    if settings.ai_provider != "groq":
        raise AiNotConfiguredError(f"Unsupported AI_PROVIDER={settings.ai_provider!r} for analysis.")
    if not settings.groq_api_key:
        raise AiNotConfiguredError("Groq analysis is enabled but GROQ_API_KEY is missing.")

    payload = {
        "model": settings.groq_model,
        "temperature": 0.1,
        "response_format": {"type": "json_object"},
        "messages": [
            {
                "role": "system",
                "content": (
                    "You are a conservative proofreading assistant for university study guides. "
                    "Return only valid JSON. Identify clear grammar, punctuation, spelling, consistency, and hyperlink issues. "
                    "Do not invent dates, readings, case names, legislation, URLs, assessments, marks, week numbers, or facts. "
                    "Suggestions must be small text replacements copied from the source where possible."
                ),
            },
            {
                "role": "user",
                "content": json.dumps(to_analysis_payload(document, hyperlinks), ensure_ascii=True),
            },
        ],
    }
    try:
        response = post_groq_chat_completion(payload)
    except AiResponseError as exc:
        fallback["warnings"].append(str(exc))
        fallback["summary"] = "Local checks completed, but Groq analysis could not run."
        return fallback
    content = response["choices"][0]["message"]["content"]
    parsed = json.loads(content)
    normalized = normalize_analysis(parsed, fallback)
    normalized["hyperlinks"] = fallback["hyperlinks"]
    return normalized


def ai_polish_document(document: StructuredDocument) -> StructuredDocument:
    if not settings.enable_ai_polish:
        return document
    if settings.ai_provider == "groq":
        try:
            return groq_polish_document(document)
        except AiResponseError:
            return document
    if settings.ai_provider == "azure":
        raise AiNotConfiguredError("Azure OpenAI is not wired for this MVP. Set AI_PROVIDER=groq.")
    raise AiNotConfiguredError(f"Unsupported AI_PROVIDER={settings.ai_provider!r}.")


def groq_polish_document(document: StructuredDocument) -> StructuredDocument:
    if not settings.groq_api_key:
        raise AiNotConfiguredError("Groq is enabled but GROQ_API_KEY is missing.")

    payload = {
        "model": settings.groq_model,
        "temperature": 0.1,
        "response_format": {"type": "json_object"},
        "messages": [
            {
                "role": "system",
                "content": (
                    "You are a conservative study guide editor. Return only valid JSON. "
                    "You may correct grammar and tidy headings, paragraphs, bullets, numbered lists, and captions. "
                    "Do not invent facts, dates, readings, URLs, legislation, case names, assessment details, or week numbers. "
                    "Do not alter table cells except for whitespace cleanup. Do not add, remove, rename, or reorder image IDs."
                ),
            },
            {
                "role": "user",
                "content": json.dumps(to_ai_payload(document), ensure_ascii=True),
            },
        ],
    }
    response = post_groq_chat_completion(payload)
    content = response["choices"][0]["message"]["content"]
    parsed = json.loads(content)
    return from_ai_payload(parsed, document)


def post_groq_chat_completion(payload: dict[str, Any]) -> dict[str, Any]:
    body = json.dumps(payload).encode("utf-8")
    request = urllib.request.Request(
        f"{settings.groq_api_base.rstrip('/')}/chat/completions",
        data=body,
        method="POST",
        headers={
            "Authorization": f"Bearer {settings.groq_api_key}",
            "Content-Type": "application/json",
            "Accept": "application/json",
            "User-Agent": "cqu-study-guide-automator/0.1",
        },
    )
    try:
        opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
        with opener.open(request, timeout=60) as response:
            return json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise AiResponseError(f"Groq request failed with HTTP {exc.code}: {detail}") from exc
    except urllib.error.URLError as exc:
        raise AiResponseError(f"Groq request failed: {exc.reason}") from exc


def deterministic_analysis(document: StructuredDocument, hyperlinks: list[dict[str, str]]) -> dict[str, Any]:
    corrections: list[dict[str, Any]] = []
    for index, block in enumerate(document.body_blocks):
        text_values = [block.text] if block.text else block.items
        for text in text_values:
            if "In Word right-click" in text:
                corrections.append(
                    {
                        "block_index": index,
                        "original": "In Word right-click",
                        "replacement": "In Word, right-click",
                        "reason": "An introductory phrase should be followed by a comma for clarity.",
                        "severity": "minor",
                    }
                )
            if "  " in text:
                corrections.append(
                    {
                        "block_index": index,
                        "original": text,
                        "replacement": " ".join(text.split()),
                        "reason": "Extra spacing can be cleaned up before publishing.",
                        "severity": "minor",
                    }
                )
    placeholder_count = sum(
        1
        for block in document.body_blocks
        if "{{" in block.text or "{%" in block.text or any("{{" in item or "{%" in item for item in block.items)
    )
    summary = "The document is ready for local conversion."
    if corrections:
        summary = f"Local checks found {len(corrections)} suggested correction{'s' if len(corrections) != 1 else ''}."
    if placeholder_count:
        summary = "The document contains template placeholders. Review these before publishing."
    return {
        "summary": summary,
        "corrections": corrections[:25],
        "hyperlinks": hyperlinks,
        "source_confidence": "high",
        "warnings": ["AI polish is disabled; only local checks were run."] if not settings.enable_ai_polish else [],
    }


def to_analysis_payload(document: StructuredDocument, hyperlinks: list[dict[str, str]]) -> dict[str, Any]:
    text_blocks: list[dict[str, Any]] = []
    for index, block in enumerate(document.body_blocks):
        if block.type in {"heading", "paragraph", "quote"} and block.text:
            text_blocks.append({"index": index, "type": block.type, "text": block.text})
        elif block.type in {"bullet_list", "numbered_list"}:
            text_blocks.append({"index": index, "type": block.type, "items": block.items})
    return {
        "required_json_shape": {
            "summary": "one sentence",
            "corrections": [
                {
                    "block_index": 0,
                    "original": "exact source text span",
                    "replacement": "corrected text span",
                    "reason": "short reason",
                    "severity": "minor|moderate|important",
                }
            ],
            "hyperlinks": [{"text": "string", "url": "string", "status": "ok|broken|needs_review"}],
            "source_confidence": "high|medium|low",
            "warnings": ["string"],
        },
        "metadata": {
            "unit_code": document.metadata.unit_code,
            "unit_name": document.metadata.unit_name,
            "week_number": document.metadata.week_number,
            "week_title": document.metadata.week_title,
            "version": document.metadata.version,
        },
        "text_blocks": text_blocks[:120],
        "hyperlinks_found_by_parser": hyperlinks,
        "instruction": "Return proofreading findings only. If there are no corrections, return an empty corrections array.",
    }


def normalize_analysis(payload: dict[str, Any], fallback: dict[str, Any]) -> dict[str, Any]:
    corrections = []
    for item in payload.get("corrections", []):
        original = clean_ai_string(item.get("original"))
        replacement = clean_ai_string(item.get("replacement"))
        if not original or not replacement or original == replacement:
            continue
        corrections.append(
            {
                "block_index": item.get("block_index", 0),
                "original": original,
                "replacement": replacement,
                "reason": clean_ai_string(item.get("reason")) or "Suggested proofreading correction.",
                "severity": clean_ai_string(item.get("severity")) or "minor",
            }
        )
    hyperlinks = payload.get("hyperlinks")
    if not isinstance(hyperlinks, list):
        hyperlinks = fallback["hyperlinks"]
    return {
        "summary": clean_ai_string(payload.get("summary")) or fallback["summary"],
        "corrections": corrections[:25],
        "hyperlinks": normalize_hyperlinks(hyperlinks),
        "source_confidence": clean_ai_string(payload.get("source_confidence")) or "medium",
        "warnings": [clean_ai_string(item) for item in payload.get("warnings", []) if clean_ai_string(item)],
    }


def normalize_hyperlinks(values: list[Any]) -> list[dict[str, str]]:
    normalized = []
    for item in values:
        if not isinstance(item, dict):
            continue
        url = clean_ai_string(item.get("url"))
        if not url:
            continue
        normalized.append(
            {
                "text": clean_ai_string(item.get("text")) or url,
                "url": url,
                "status": clean_ai_string(item.get("status")) or "needs_review",
                "detail": clean_ai_string(item.get("detail")),
                "http_status": clean_ai_string(item.get("http_status")),
            }
        )
    return normalized


def to_ai_payload(document: StructuredDocument) -> dict[str, Any]:
    return {
        "schema": {
            "unit_code": "string",
            "unit_name": "string",
            "week_number": "string",
            "week_title": "string",
            "version": "string",
            "body_blocks": [
                {
                    "type": "heading|paragraph|bullet_list|numbered_list|table|image|quote",
                    "level": "number",
                    "number": "string",
                    "text": "string",
                    "items": ["string"],
                    "rows": [["string"]],
                    "image_id": "string",
                    "caption": "string",
                }
            ],
            "validation_notes": {
                "possible_hallucinations": [],
                "missing_images": [],
                "source_confidence": "high|medium|low",
            },
        },
        "unit_code": document.metadata.unit_code,
        "unit_name": document.metadata.unit_name,
        "week_number": document.metadata.week_number,
        "week_title": document.metadata.week_title,
        "version": document.metadata.version,
        "body_blocks": [block_to_dict(block) for block in document.body_blocks],
        "available_image_ids": sorted(document.images),
        "instruction": "Return the same JSON shape with polished text only.",
    }


def block_to_dict(block: ContentBlock) -> dict[str, Any]:
    return {
        "type": block.type,
        "level": block.level,
        "number": block.number,
        "text": block.text,
        "items": block.items,
        "rows": block.rows,
        "image_id": block.image_id,
        "caption": block.caption,
    }


def from_ai_payload(payload: dict[str, Any], original: StructuredDocument) -> StructuredDocument:
    blocks = [block_from_dict(item) for item in payload.get("body_blocks", [])]
    if not blocks:
        raise AiResponseError("Groq returned no body_blocks.")
    validate_ai_blocks(blocks, original)
    metadata = original.metadata
    metadata.unit_code = clean_ai_string(payload.get("unit_code")) or metadata.unit_code
    metadata.unit_name = clean_ai_string(payload.get("unit_name")) or metadata.unit_name
    metadata.week_number = clean_ai_string(payload.get("week_number")) or metadata.week_number
    metadata.week_title = clean_ai_string(payload.get("week_title")) or metadata.week_title
    metadata.version = clean_ai_string(payload.get("version")) or metadata.version
    return StructuredDocument(metadata=metadata, body_blocks=blocks, images=original.images)


def block_from_dict(item: dict[str, Any]) -> ContentBlock:
    block_type = item.get("type")
    if block_type not in {"heading", "paragraph", "bullet_list", "numbered_list", "table", "image", "quote"}:
        raise AiResponseError(f"Unsupported block type from Groq: {block_type!r}")
    return ContentBlock(
        type=block_type,
        text=clean_ai_string(item.get("text")),
        level=int(item.get("level") or 1),
        number=clean_ai_string(item.get("number")),
        items=[clean_ai_string(value) for value in item.get("items", [])],
        rows=[[clean_ai_string(cell) for cell in row] for row in item.get("rows", [])],
        image_id=clean_ai_string(item.get("image_id")),
        caption=clean_ai_string(item.get("caption")),
    )


def validate_ai_blocks(blocks: list[ContentBlock], original: StructuredDocument) -> None:
    original_image_ids = [block.image_id for block in original.body_blocks if block.type == "image"]
    returned_image_ids = [block.image_id for block in blocks if block.type == "image"]
    if returned_image_ids != original_image_ids:
        raise AiResponseError("Groq response changed image IDs or image order.")

    original_tables = [block.rows for block in original.body_blocks if block.type == "table"]
    returned_tables = [block.rows for block in blocks if block.type == "table"]
    if len(returned_tables) != len(original_tables):
        raise AiResponseError("Groq response changed the number of tables.")
    for original_rows, returned_rows in zip(original_tables, returned_tables, strict=True):
        if table_shape(original_rows) != table_shape(returned_rows):
            raise AiResponseError("Groq response changed a table shape.")


def table_shape(rows: list[list[str]]) -> list[int]:
    return [len(row) for row in rows]


def clean_ai_string(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()
