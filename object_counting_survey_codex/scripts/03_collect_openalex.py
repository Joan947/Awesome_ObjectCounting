"""Collect OpenAlex metadata for query-bank terms.

Semantic Scholar is intentionally skipped until an API key is available. This
script uses OpenAlex's public API and the user's polite-pool mailto address.
"""

from __future__ import annotations

import csv
import json
import os
import re
import time
import urllib.parse
import urllib.request
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
QUERY_BANK = ROOT / "query_bank.yaml"
OUT = ROOT / "data" / "raw_api_results" / "openalex_results.csv"
MAILTO = os.environ.get("OPENALEX_MAILTO", "jowusu1@uwyo.edu")
API = "https://api.openalex.org/works"


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


def inverted_abstract(index: dict | None) -> str:
    if not index:
        return ""
    words = []
    for word, positions in index.items():
        for pos in positions:
            words.append((pos, word))
    return " ".join(word for _, word in sorted(words))


def clean_doi(value: str | None) -> str:
    if not value:
        return ""
    return re.sub(r"^https?://(dx\.)?doi\.org/", "", value, flags=re.I).lower()


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
                "search": term,
                "filter": "from_publication_date:2015-01-01",
                "per-page": 50,
                "mailto": MAILTO,
            }
        )
        url = f"{API}?{params}"
        try:
            data = get_json(url)
        except Exception as exc:
            rows.append({"source": "OpenAlex", "source_query": term, "query_bucket": bucket, "error": str(exc)})
            continue
        for item in data.get("results", []):
            title = item.get("title") or ""
            key = item.get("id") or title.lower()
            if key in seen:
                continue
            seen.add(key)
            authors = "; ".join(
                (auth.get("author") or {}).get("display_name", "")
                for auth in item.get("authorships", [])
                if (auth.get("author") or {}).get("display_name")
            )
            rows.append(
                {
                    "record_id": f"OA_{len(rows)+1:06d}",
                    "source": "OpenAlex",
                    "source_query": term,
                    "query_bucket": bucket,
                    "paper_title": title,
                    "authors": authors,
                    "year": item.get("publication_year") or "",
                    "venue": ((item.get("primary_location") or {}).get("source") or {}).get("display_name", ""),
                    "doi": clean_doi(item.get("doi")),
                    "arxiv_id": "",
                    "paper_url": item.get("id") or "",
                    "abstract": inverted_abstract(item.get("abstract_inverted_index")),
                    "author_keywords": "",
                    "index_keywords": "; ".join(c.get("display_name", "") for c in item.get("concepts", [])[:10]),
                    "citation_count": item.get("cited_by_count") or "",
                    "document_type": item.get("type") or "",
                    "dataset_or_benchmark": "",
                    "modality": "",
                    "task_category": "API expansion",
                    "method_family": "",
                    "input_type": "",
                    "output_type": "",
                    "application_area": "",
                    "relevance_score": "",
                    "tier": "",
                    "reason_kept_or_removed": "Collected from OpenAlex query.",
                    "openalex_id": item.get("id") or "",
                    "error": "",
                }
            )
        time.sleep(0.3)
    fields = sorted({key for row in rows for key in row})
    with OUT.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)
    print(f"Wrote {len(rows)} OpenAlex rows to {OUT}")


if __name__ == "__main__":
    main()
