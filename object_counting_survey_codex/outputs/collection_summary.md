# Collection Summary

Generated: 2026-05-25

## Day 2 Status

Semantic Scholar expansion is deferred until an API key is available. Current Day 2 processing used:

- Google Scholar seed workbook: 80 unique seed records
- Scopus exports: 3,287 imported records from 10 CSV files
- OpenAlex: 1,433 raw rows using `mailto=jowusu1@uwyo.edu`
- Crossref: 634 raw rows
- arXiv: 22 raw rows from a reduced high-yield query subset
- Papers With Code: attempted, but current responses were non-JSON/error rows and no usable title records were merged

## Current Outputs

- `data/master_papers_raw.csv`: 5,437 raw records before deduplication/screening
- `data/literature_matrix_clean.xlsx`: 700 included Day 2 candidates
- `data/screening_log.xlsx`: 700 included candidates plus 4,737 excluded/duplicate/overflow records
- `outputs/core_reading_set.xlsx`: 150 A-core automated candidates
- `data/raw_api_results/scopus_imported_records.csv`: normalized Scopus import

## Candidate Tiers

- A_core: 150
- B_important: 200
- C_background: 350

## Notes

The 700-record literature matrix is an automated Day 2 working corpus, not the final screened corpus. The full raw record set is preserved in `data/master_papers_raw.csv`, and excluded/overflow records remain in `data/screening_log.xlsx` for manual rescue during gap filling.
