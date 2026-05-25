"""Resolve DOI and publisher metadata with Crossref."""

from __future__ import annotations

import csv
import json
import time
import urllib.parse
import urllib.request
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
QUERY_BANK = ROOT / "query_bank.yaml"
OUT = ROOT / "data" / "raw_api_results" / "crossref_results.csv"
API = "https://api.crossref.org/works"
MAILTO = "jowusu1@uwyo.edu"


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
    return terms


def first(value) -> str:
    if isinstance(value, list) and value:
        return str(value[0])
    return "" if value is None else str(value)


def get_json(url: str) -> dict:
    req = urllib.request.Request(url, headers={"User-Agent": f"object-counting-survey/0.1 (mailto:{MAILTO})"})
    with urllib.request.urlopen(req, timeout=45) as response:
        return json.loads(response.read().decode("utf-8"))


def main() -> None:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    rows = []
    seen = set()
    for bucket, term in query_terms():
        params = urllib.parse.urlencode(
            {
                "query.title": term,
                "filter": "from-pub-date:2015-01-01",
                "rows": 25,
                "mailto": MAILTO,
            }
        )
        try:
            data = get_json(f"{API}?{params}")
        except Exception as exc:
            rows.append({"source": "Crossref", "source_query": term, "query_bucket": bucket, "error": str(exc)})
            continue
        for item in data.get("message", {}).get("items", []):
            title = first(item.get("title"))
            doi = (item.get("DOI") or "").lower()
            key = doi or title.lower()
            if not key or key in seen:
                continue
            seen.add(key)
            authors = []
            for author in item.get("author", []):
                authors.append(" ".join(part for part in [author.get("given"), author.get("family")] if part))
            year = ""
            date_parts = item.get("published-print", item.get("published-online", item.get("created", {}))).get("date-parts", [])
            if date_parts and date_parts[0]:
                year = date_parts[0][0]
            rows.append(
                {
                    "record_id": f"CR_{len(rows)+1:06d}",
                    "source": "Crossref",
                    "source_query": term,
                    "query_bucket": bucket,
                    "paper_title": title,
                    "authors": "; ".join(authors),
                    "year": year,
                    "venue": first(item.get("container-title")),
                    "doi": doi,
                    "arxiv_id": "",
                    "paper_url": item.get("URL") or "",
                    "abstract": item.get("abstract") or "",
                    "author_keywords": "",
                    "index_keywords": "",
                    "citation_count": item.get("is-referenced-by-count") or "",
                    "document_type": item.get("type") or "",
                    "dataset_or_benchmark": "",
                    "modality": "",
                    "task_category": "DOI/metadata expansion",
                    "method_family": "",
                    "input_type": "",
                    "output_type": "",
                    "application_area": "",
                    "relevance_score": "",
                    "tier": "",
                    "reason_kept_or_removed": "Collected from Crossref query.",
                    "publisher": item.get("publisher") or "",
                    "error": "",
                }
            )
        time.sleep(0.3)
    fields = sorted({key for row in rows for key in row})
    with OUT.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)
    print(f"Wrote {len(rows)} Crossref rows to {OUT}")


if __name__ == "__main__":
    main()
