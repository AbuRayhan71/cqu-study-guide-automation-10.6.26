from __future__ import annotations

import ssl
import urllib.error
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

from docx import Document
from docx.oxml.ns import qn


CHECKABLE_SCHEMES = {"http", "https"}
DEFAULT_TIMEOUT_SECONDS = 8
MAX_WORKERS = 8


@dataclass(frozen=True)
class Hyperlink:
    text: str
    url: str


def extract_hyperlinks(path: Path) -> list[dict[str, str]]:
    document = Document(path)
    relationships = document.part.rels
    links: list[Hyperlink] = []

    for node in document.part._element.xpath(".//w:hyperlink[@r:id]"):
        rel_id = node.get(qn("r:id"))
        if not rel_id or rel_id not in relationships:
            continue
        relationship = relationships[rel_id]
        if "hyperlink" not in relationship.reltype:
            continue
        url = relationship.target_ref
        text = "".join(node.xpath(".//w:t/text()")).strip() or url
        links.append(Hyperlink(text=text, url=url))

    known_urls = {link.url for link in links}
    for relationship in relationships.values():
        if "hyperlink" not in relationship.reltype:
            continue
        url = relationship.target_ref
        if url in known_urls:
            continue
        links.append(Hyperlink(text=url, url=url))
        known_urls.add(url)

    return [link_to_dict(link, "present", "Hyperlink was found in the DOCX.") for link in dedupe_links(links)]


def check_hyperlinks(links: list[dict[str, str]]) -> list[dict[str, str]]:
    if not links:
        return []

    checked: list[dict[str, str] | None] = [None] * len(links)
    with ThreadPoolExecutor(max_workers=min(MAX_WORKERS, len(links))) as executor:
        future_map = {
            executor.submit(check_hyperlink, link, DEFAULT_TIMEOUT_SECONDS): index
            for index, link in enumerate(links)
        }
        for future in as_completed(future_map):
            index = future_map[future]
            try:
                checked[index] = future.result()
            except Exception as exc:
                original = links[index]
                checked[index] = {
                    "text": original.get("text") or original.get("url", ""),
                    "url": original.get("url", ""),
                    "status": "needs_review",
                    "detail": f"Link check failed unexpectedly: {exc}",
                }

    return [item for item in checked if item is not None]


def check_hyperlink(link: dict[str, str], timeout_seconds: int = DEFAULT_TIMEOUT_SECONDS) -> dict[str, str]:
    url = (link.get("url") or "").strip()
    text = (link.get("text") or url).strip()
    if not url:
        return {"text": text, "url": url, "status": "broken", "detail": "Missing URL."}

    parsed = urlparse(url)
    if not parsed.scheme:
        return {
            "text": text,
            "url": url,
            "status": "needs_review",
            "detail": "Relative or internal link; verify manually.",
        }
    if parsed.scheme.lower() not in CHECKABLE_SCHEMES:
        return {
            "text": text,
            "url": url,
            "status": "needs_review",
            "detail": f"{parsed.scheme} links cannot be checked over HTTP.",
        }

    response = fetch_url_status(url, timeout_seconds)
    status = classify_status(response["status_code"], response.get("error", ""))
    return {
        "text": text,
        "url": url,
        "status": status,
        "detail": response["detail"],
        "http_status": str(response["status_code"]) if response["status_code"] else "",
    }


def fetch_url_status(url: str, timeout_seconds: int) -> dict[str, Any]:
    head = request_url(url, "HEAD", timeout_seconds)
    if should_retry_with_get(head):
        return request_url(url, "GET", timeout_seconds)
    return head


def request_url(url: str, method: str, timeout_seconds: int) -> dict[str, Any]:
    request = urllib.request.Request(
        url,
        method=method,
        headers={
            "Accept": "*/*",
            "User-Agent": "cqu-study-guide-automator/0.1",
        },
    )
    try:
        with urllib.request.urlopen(request, timeout=timeout_seconds) as response:
            code = int(response.status)
            return {
                "status_code": code,
                "detail": f"HTTP {code} via {method}.",
            }
    except urllib.error.HTTPError as exc:
        return {
            "status_code": int(exc.code),
            "detail": f"HTTP {exc.code} via {method}.",
        }
    except urllib.error.URLError as exc:
        reason = getattr(exc, "reason", exc)
        return {
            "status_code": 0,
            "error": str(reason),
            "detail": f"Could not reach URL: {reason}",
        }
    except TimeoutError:
        return {"status_code": 0, "error": "timeout", "detail": "Link check timed out."}
    except ssl.SSLError as exc:
        return {"status_code": 0, "error": str(exc), "detail": f"TLS check failed: {exc}"}


def should_retry_with_get(result: dict[str, Any]) -> bool:
    return result.get("status_code") in {403, 405}


def classify_status(status_code: int, error: str = "") -> str:
    if 200 <= status_code < 400:
        return "ok"
    if status_code in {401, 403, 429}:
        return "needs_review"
    if status_code == 0 and error:
        return "needs_review"
    if status_code >= 400:
        return "broken"
    return "needs_review"


def dedupe_links(links: list[Hyperlink]) -> list[Hyperlink]:
    seen: set[tuple[str, str]] = set()
    unique = []
    for link in links:
        key = (link.text, link.url)
        if key in seen:
            continue
        seen.add(key)
        unique.append(link)
    return unique


def link_to_dict(link: Hyperlink, status: str, detail: str) -> dict[str, str]:
    return {"text": link.text, "url": link.url, "status": status, "detail": detail}
