# Object Counting Survey Codex Workspace

This folder follows the final 3-day collection guide in `../final_plan/object_counting_survey_final_3day_guide.pdf`.

## Purpose

Build a reproducible literature collection pipeline for:

> Object Counting Across Modalities: A Survey of Image, Video, Open-Vocabulary, and Emerging Visual Counting Methods.

The workflow combines supervised Google Scholar seed discovery, Scopus exports, legitimate scholarly APIs, deduplication, screening, BibTeX validation, and final survey organization.

## Day 1

1. Continue seed collection in `data/seed_papers_google_scholar.xlsx`.
2. Record every Scholar and Scopus query in `data/query_log.xlsx`.
3. Save Scopus CSV or BibTeX exports in `data/scopus_exports/`.
4. Do not scrape Google Scholar. Use Scholar only for supervised manual seed discovery.

## Day 2

1. Import Scopus exports with `scripts/00_import_scopus_exports.py`.
2. Collect and expand records from legitimate APIs:
   - Papers With Code
   - Semantic Scholar
   - OpenAlex
   - Crossref
   - arXiv
3. Save raw API outputs in `data/raw_api_results/`.
4. Merge and deduplicate into `data/literature_matrix_clean.xlsx`.

## Day 3

1. Screen uncertain records and fill modality gaps.
2. Generate the core reading set, taxonomy groups, benchmark compendium, collection summary, PRISMA-style screening summary, final outline, and validated BibTeX.
3. Manually review all A-tier papers and questionable B-tier papers before writing.

## Current Seed Source

The starting seed workbook was copied from `../Object_Counting.xlsx` into:

`data/seed_papers_google_scholar.xlsx`

## Ethical Access Rules

- Do not scrape Google Scholar.
- Stop if Scholar slows down, blocks, or asks for CAPTCHA.
- Use Scopus only through legitimate institutional access and built-in export tools.
- Use APIs according to their terms and rate limits.
