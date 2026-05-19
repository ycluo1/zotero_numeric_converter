@echo off
REM Put paper.docx, references.bib, and zotero.rdf in this folder first.
python zotero_numeric_converter.py --docx paper.docx --bib references.bib --rdf zotero.rdf --out paper_zotero.docx --dry-run
pause
