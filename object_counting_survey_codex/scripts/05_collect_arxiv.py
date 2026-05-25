"""Collect recent/gap-fill preprints from the arXiv API."""

from __future__ import annotations

import csv
import re
import time
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
QUERY_BANK = ROOT / "query_bank.yaml"
OUT = ROOT / "data" / "raw_api_results" / "arxiv_results.csv"
API = "https://export.arxiv.org/api/query"
NS = {"atom": "http://www.w3.org/2005/Atom", "arxiv": "http://arxiv.org/schemas/atom"}


def query_terms() -> list[tuple[str, str]]:
    terms: list[tuple[str, str]] = []
    current = "general"
    for line in QUERY_BANK.read_text(encoding="utf-8").splitlines():
        if not line.strip() or line.strip().startswith("#"):
            continue
        if not line.startswith(" ") and line.endswith(":"):
            current = line.rstrip(":").strip()
        elif line.strip().startswith("- "):
            terms.append((current, line.strip()[2:].strip()))
    preferred = {
        "few-shot object counting",
        "exemplar-based object counting",
        "class-agnostic counting",
        "zero-shot object counting",
        "open-vocabulary object counting",
        "text-guided object counting",
        "SAM object counting",
        "CLIP object counting",
        "video object counting",
        "unique instance counting video",
        "3D object counting",
        "point cloud object counting",
        "remote sensing object counting",
        "aerial image object counting",
        "cell counting computer vision",
        "microscopy counting",
        "FSC-147",
        "CountBench",
        "OmniCount",
        "VideoCount",
    }
    selected = [item for item in terms if item[1] in preferred]
    return selected or terms[:20]


def text(node, path: str) -> str:
    found = node.find(path, NS)
    return re.sub(r"\s+", " ", found.text).strip() if found is not None and found.text else ""


def arxiv_id(url: str) -> str:
    match = re.search(r"abs/([^v]+(?:\.\d+)?)", url)
    return match.group(1) if match else ""


def main() -> None:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    rows = []
    seen = set()
    for bucket, term in query_terms():
        search = f'all:"{term}"'
        params = urllib.parse.urlencode(
            {
                "search_query": search,
                "start": 0,
                "max_results": 10,
                "sortBy": "submittedDate",
                "sortOrder": "descending",
            }
        )
        url = f"{API}?{params}"
        try:
            with urllib.request.urlopen(url, timeout=15) as response:
                root = ET.fromstring(response.read())
        except Exception as exc:
            rows.append({"source": "arXiv", "source_query": term, "query_bucket": bucket, "error": str(exc)})
            continue
        for entry in root.findall("atom:entry", NS):
            title = text(entry, "atom:title")
            link = text(entry, "atom:id")
            aid = arxiv_id(link)
            if aid in seen:
                continue
            seen.add(aid)
            authors = "; ".join(text(author, "atom:name") for author in entry.findall("atom:author", NS))
            year = text(entry, "atom:published")[:4]
            doi_node = entry.find("arxiv:doi", NS)
            rows.append(
                {
                    "record_id": f"ARXIV_{len(rows)+1:06d}",
                    "source": "arXiv",
                    "source_query": term,
                    "query_bucket": bucket,
                    "paper_title": title,
                    "authors": authors,
                    "year": year,
                    "venue": "arXiv",
                    "doi": doi_node.text if doi_node is not None and doi_node.text else "",
                    "arxiv_id": aid,
                    "paper_url": link,
                    "abstract": text(entry, "atom:summary"),
                    "author_keywords": "",
                    "index_keywords": "; ".join(cat.get("term", "") for cat in entry.findall("atom:category", NS)),
                    "citation_count": "",
                    "document_type": "preprint",
                    "dataset_or_benchmark": "",
                    "modality": "",
                    "task_category": "Recent/gap-fill preprint",
                    "method_family": "",
                    "input_type": "",
                    "output_type": "",
                    "application_area": "",
                    "relevance_score": "",
                    "tier": "",
                    "reason_kept_or_removed": "Collected from arXiv query.",
                    "error": "",
                }
            )
        time.sleep(1.0)
    fields = sorted({key for row in rows for key in row})
    with OUT.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)
    print(f"Wrote {len(rows)} arXiv rows to {OUT}")


if __name__ == "__main__":
    main()
