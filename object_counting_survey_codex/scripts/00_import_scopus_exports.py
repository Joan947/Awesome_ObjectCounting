"""Import Scopus CSV exports into a normalized project schema."""

from __future__ import annotations

import csv
import re
from collections import Counter, defaultdict
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "data"
SCOPUS_DIR = DATA / "scopus_exports"
OUT_DIR = DATA / "raw_api_results"
OUT_CSV = OUT_DIR / "scopus_imported_records.csv"

DATASETS = [
    "FSC-147",
    "FSCD-147",
    "CountBench",
    "OmniCount",
    "CARPK",
    "ShanghaiTech",
    "TAO-Count",
    "VideoCount",
    "CountVid",
    "MovingDroneCrowd",
    "DroneCrowd",
]


def clean_text(value: object) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def norm_doi(value: object) -> str:
    text = clean_text(value)
    text = re.sub(r"^https?://(dx\.)?doi\.org/", "", text, flags=re.I)
    return text.rstrip("}. ").lower()


def bucket_from_filename(name: str) -> str:
    stem = Path(name).stem.removeprefix("scopus_")
    return stem.replace("_", " ")


def joined_text(row: dict[str, str]) -> str:
    parts = [
        row.get("Title", ""),
        row.get("Abstract", ""),
        row.get("Author Keywords", ""),
        row.get("Index Keywords", ""),
        row.get("Source title", ""),
    ]
    return " ".join(clean_text(p) for p in parts).lower()


def extract_datasets(text: str) -> str:
    found = []
    for dataset in DATASETS:
        if dataset.lower() in text.lower():
            found.append(dataset)
    return "; ".join(dict.fromkeys(found))


def infer_modality(row: dict[str, str], export_filename: str) -> str:
    text = joined_text(row)
    name = export_filename.lower()
    if "medical" in name or "microscopy" in name or any(k in text for k in ["cell counting", "microscopy", "bacteria", "histology"]):
        return "medical; microscopy"
    if "remote" in name or any(k in text for k in ["remote sensing", "aerial", "uav", "drone", "satellite"]):
        return "remote_sensing; aerial"
    if "3d" in name or any(k in text for k in ["3d object", "point cloud", "rgb-d", "depth", "lidar"]):
        return "3d; point_cloud"
    if "video" in name or any(k in text for k in ["video", "tracking", "tracklet", "temporal"]):
        return "video"
    if any(k in text for k in ["thermal", "infrared"]):
        return "thermal"
    if any(k in text for k in ["event camera", "event-based", "neuromorphic"]):
        return "event_camera"
    if any(k in text for k in ["vision-language", "vision language", "text-guided", "open-vocabulary", "open vocabulary", "clip", "sam"]):
        return "image; multimodal"
    return "image"


def infer_task_category(row: dict[str, str], export_filename: str) -> str:
    text = joined_text(row)
    name = export_filename.lower()
    labels = []
    if "benchmark" in name or "dataset" in name or any(k.lower() in text for k in DATASETS) or any(k in text for k in ["benchmark", "dataset"]):
        labels.append("Dataset/Benchmark")
    if "fewshot" in name or any(k in text for k in ["few-shot", "few shot", "exemplar", "class-agnostic", "class agnostic"]):
        labels.append("Few-shot/Class-agnostic")
    if "openvocab" in name or any(k in text for k in ["zero-shot", "zero shot", "open-vocabulary", "open vocabulary", "text-guided"]):
        labels.append("Zero-shot/Open-vocabulary")
    if any(k in text for k in ["foundation model", "vision-language", "clip", "sam"]):
        labels.append("Foundation/Vision-language")
    if any(k in text for k in ["density map", "density estimation", "crowd counting"]):
        labels.append("Density-based")
    if any(k in text for k in ["detect", "bounding box", "yolo", "instance segmentation"]):
        labels.append("Detection/Segmentation-based")
    if any(k in text for k in ["tracking", "re-identification", "tracklet", "unique instance"]):
        labels.append("Tracking/Unique-count")
    if not labels:
        labels.append("Application-specific")
    return "; ".join(dict.fromkeys(labels))


def infer_input_type(modality: str, task: str, row: dict[str, str]) -> str:
    text = joined_text(row)
    inputs = []
    if "video" in modality:
        inputs.append("Video")
    if "point_cloud" in modality:
        inputs.append("Point Cloud")
    if "remote_sensing" in modality:
        inputs.append("Remote Sensing Image")
    if "medical" in modality or "microscopy" in modality:
        inputs.append("Medical/Microscopy Image")
    if "thermal" in modality:
        inputs.append("Thermal Image/Video")
    if "event_camera" in modality:
        inputs.append("Event Camera")
    if not inputs:
        inputs.append("Image")
    if any(k in text for k in ["text-guided", "open-vocabulary", "prompt", "language"]):
        inputs.append("Text Prompt")
    if any(k in text for k in ["exemplar", "few-shot", "few shot", "reference"]):
        inputs.append("Exemplar Boxes/Patches")
    if "rgb-d" in text or "depth" in text:
        inputs.append("Depth/RGB-D")
    return "; ".join(dict.fromkeys(inputs))


def infer_output_type(row: dict[str, str]) -> str:
    text = joined_text(row)
    outputs = []
    if any(k in text for k in ["density map", "density estimation"]):
        outputs.append("Density Map + Count")
    if any(k in text for k in ["track", "trajectory", "unique instance"]):
        outputs.append("Tracks + Unique Count")
    if any(k in text for k in ["bounding box", "detect", "localization"]):
        outputs.append("Boxes + Count")
    if any(k in text for k in ["mask", "segmentation"]):
        outputs.append("Masks + Count")
    if not outputs:
        outputs.append("Count Only")
    return "; ".join(dict.fromkeys(outputs))


def infer_application_area(modality: str, row: dict[str, str]) -> str:
    text = joined_text(row)
    if "medical" in modality or any(k in text for k in ["cell", "bacteria", "histology", "tumor", "microscopy"]):
        return "medical/microscopy"
    if "remote_sensing" in modality:
        return "remote sensing/aerial"
    if any(k in text for k in ["crop", "cotton", "fruit", "plant", "agricultur"]):
        return "agriculture"
    if any(k in text for k in ["crowd", "pedestrian", "traffic", "vehicle"]):
        return "crowd/traffic"
    if "point_cloud" in modality:
        return "autonomous driving/3D"
    return "general visual counting"


def map_row(row: dict[str, str], export_filename: str, row_index: int) -> dict[str, str]:
    title = clean_text(row.get("Title"))
    modality = infer_modality(row, export_filename)
    task = infer_task_category(row, export_filename)
    text = joined_text(row)
    record = {
        "record_id": f"SCOPUS_{Path(export_filename).stem}_{row_index:05d}",
        "source": "Scopus",
        "source_query": bucket_from_filename(export_filename),
        "export_filename": export_filename,
        "paper_title": title,
        "authors": clean_text(row.get("Authors")),
        "author_full_names": clean_text(row.get("Author full names")),
        "year": clean_text(row.get("Year")),
        "venue": clean_text(row.get("Source title")),
        "doi": norm_doi(row.get("DOI")),
        "arxiv_id": "",
        "paper_url": clean_text(row.get("Link")),
        "abstract": clean_text(row.get("Abstract")),
        "author_keywords": clean_text(row.get("Author Keywords")),
        "index_keywords": clean_text(row.get("Index Keywords")),
        "citation_count": clean_text(row.get("Cited by")),
        "document_type": clean_text(row.get("Document Type")),
        "publication_stage": clean_text(row.get("Publication Stage")),
        "open_access": clean_text(row.get("Open Access")),
        "affiliations": clean_text(row.get("Affiliations")),
        "references": clean_text(row.get("References")),
        "eid": clean_text(row.get("EID")),
        "issn": clean_text(row.get("ISSN")),
        "isbn": clean_text(row.get("ISBN")),
        "pubmed_id": clean_text(row.get("PubMed ID")),
        "publisher": clean_text(row.get("Publisher")),
        "dataset_or_benchmark": extract_datasets(text),
        "modality": modality,
        "task_category": task,
        "method_family": task,
        "input_type": infer_input_type(modality, task, row),
        "output_type": infer_output_type(row),
        "application_area": infer_application_area(modality, row),
        "relevance_score": "",
        "tier": "",
        "reason_kept_or_removed": "",
    }
    for key, value in row.items():
        safe_key = re.sub(r"[^a-z0-9]+", "_", key.lower()).strip("_")
        record[f"scopus_{safe_key}"] = clean_text(value)
    return record


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    records: list[dict[str, str]] = []
    by_file: Counter[str] = Counter()
    by_year: Counter[str] = Counter()
    by_doc_type: Counter[str] = Counter()
    by_modality: Counter[str] = Counter()
    for path in sorted(SCOPUS_DIR.glob("*.csv")):
        with path.open(newline="", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f)
            for i, row in enumerate(reader, start=1):
                record = map_row(row, path.name, i)
                records.append(record)
                by_file[path.name] += 1
                by_year[record["year"]] += 1
                by_doc_type[record["document_type"]] += 1
                by_modality[record["modality"]] += 1
    if not records:
        raise SystemExit(f"No Scopus CSV files found in {SCOPUS_DIR}")
    field_order = []
    seen = set()
    for record in records:
        for key in record:
            if key not in seen:
                field_order.append(key)
                seen.add(key)
    with OUT_CSV.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=field_order, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(records)
    print(f"Wrote {len(records)} Scopus rows to {OUT_CSV}")
    print("\nCounts by export file:")
    for key, value in by_file.most_common():
        print(f"  {key}: {value}")
    print("\nCounts by year:")
    for key, value in sorted(by_year.items(), reverse=True):
        print(f"  {key or 'UNKNOWN'}: {value}")
    print("\nCounts by document type:")
    for key, value in by_doc_type.most_common():
        print(f"  {key or 'UNKNOWN'}: {value}")
    print("\nPreliminary modality counts:")
    for key, value in by_modality.most_common():
        print(f"  {key}: {value}")


if __name__ == "__main__":
    main()
