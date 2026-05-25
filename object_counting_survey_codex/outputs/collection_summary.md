# Collection Summary

Generated: 2026-05-25

## Day 2 Status

Day 2 automated collection now includes the Semantic Scholar expansion. Current Day 2 processing used:

- Google Scholar seed workbook: 80 unique seed records
- Scopus exports: 3,287 imported records from 10 CSV files
- Semantic Scholar: 620 enriched metadata rows, 624 query-bank search rows, 1,539 reference rows, and 1,392 citation rows
- OpenAlex: 1,433 raw rows using `mailto=jowusu1@uwyo.edu`
- Crossref: 634 raw rows
- arXiv: 22 raw rows from a reduced high-yield query subset
- Papers With Code: attempted, but current responses were non-JSON/error rows and no usable title records were merged

## Current Outputs

- `data/master_papers_raw.csv`: 18,919 raw records before deduplication/screening
- `data/raw_api_results/deduplicated_unique_records.csv`: 5,657 unique records after deduplication
- `data/raw_api_results/duplicate_removed_records.csv`: 13,262 duplicate records removed
- `data/literature_matrix_clean.xlsx`: 700 included Day 2 candidates
- `data/screening_log.xlsx`: 700 included candidates plus 18,219 excluded/duplicate/overflow records
- `outputs/core_reading_set.xlsx`: 150 A-core automated candidates
- `data/raw_api_results/scopus_imported_records.csv`: normalized Scopus import, kept locally and ignored by git because of size
- `data/raw_api_results/semantic_scholar_enriched_records.csv`: Semantic Scholar batch/title metadata enrichment
- `data/raw_api_results/semantic_scholar_search_results.csv`: Semantic Scholar query-bank topic search records
- `data/raw_api_results/semantic_scholar_references.csv`: backward citation expansion for selected seed papers
- `data/raw_api_results/semantic_scholar_citations.csv`: forward citation expansion for selected seed papers

## Deduplication Snapshot

- Raw records: 18,919
- Records after deduplication: 5,657
- Records removed as duplicates: 13,262
- Records with DOI after deduplication: 5,365
- Records missing DOI after deduplication: 292
- Records with abstract after deduplication: 4,781
- Records missing abstract after deduplication: 876

## Candidate Tiers

- A_core: 150
- B_important: 200
- C_background: 350

## Notes

The 700-record literature matrix is an automated Day 2 working corpus, not the final screened corpus. The full raw record set is preserved in `data/master_papers_raw.csv`, and excluded/overflow records remain in `data/screening_log.xlsx` for manual rescue during gap filling.
