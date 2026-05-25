"""Collect benchmark/task records from Papers With Code.

This script avoids Google Scholar entirely and uses the public Papers With
Code API. It is intentionally conservative: it queries the local query bank,
stores raw metadata, and leaves deduplication to 06_merge_deduplicate.py.
"""

from __future__ import annotations

import csv
import json
import time
import urllib.parse
import urllib.request
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "data" / "raw_api_results" / "paperswithcode_results.csv"
QUERY_BANK = ROOT / "query_bank.yaml"
API = "https://paperswithcode.com/api/v1/papers/"


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


def get_json(url: str) -> dict:
    req = urllib.request.Request(url, headers={"User-Agent": "object-counting-survey/0.1"})
    with urllib.request.urlopen(req, timeout=30) as response:
        return json.loads(response.read().decode("utf-8"))


def main() -> None:
    OUT.parent.mkdir(parents=True, exist_ok=True)
    rows = []
    seen = set()
    for bucket, term in query_terms():
        params = urllib.parse.urlencode({"q": term, "page_size": 50})
        url = f"{API}?{params}"
        try:
            data = get_json(url)
        except Exception as exc:
            rows.append({"source": "Papers With Code", "source_query": term, "error": str(exc)})
            continue
        for item in data.get("results", []):
            title = item.get("title") or ""
            key = (title.lower(), item.get("arxiv_id") or item.get("url_abs") or item.get("paper_url"))
            if key in seen:
                continue
            seen.add(key)
            rows.append(
                {
                    "record_id": f"PWC_{len(rows)+1:06d}",
                    "source": "Papers With Code",
                    "source_query": term,
                    "query_bucket": bucket,
                    "paper_title": title,
                    "authors": "",
                    "year": "",
                    "venue": "",
                    "doi": "",
                    "arxiv_id": item.get("arxiv_id") or "",
                    "paper_url": item.get("url_abs") or item.get("paper_url") or "",
                    "abstract": item.get("abstract") or "",
                    "author_keywords": "",
                    "index_keywords": "",
                    "citation_count": "",
                    "document_type": "",
                    "dataset_or_benchmark": "",
                    "modality": "",
                    "task_category": "Benchmark/API expansion",
                    "method_family": "",
                    "input_type": "",
                    "output_type": "",
                    "application_area": "",
                    "relevance_score": "",
                    "tier": "",
                    "reason_kept_or_removed": "Collected from Papers With Code query.",
                    "github_url": item.get("url_pdf") or "",
                    "api_url": item.get("url") or "",
                    "error": "",
                }
            )
        time.sleep(0.5)
    fields = sorted({key for row in rows for key in row})
    with OUT.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)
    print(f"Wrote {len(rows)} Papers With Code rows to {OUT}")


if __name__ == "__main__":
    main()
