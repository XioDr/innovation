# innovation

Tooling and deliverables for a **Persian-language (RTL) technical and feasibility proposal** for the "Hamrah-e Entekhab" decision-support platform.

## Contents

| File | Description |
| --- | --- |
| `build_proposal.py` | Generates the expanded proposal as a `.docx` with python-docx (A4, RTL layout, Persian typography, branded palette, tables and charts rendered with Pillow). |
| `fix_word_rtl.py` | Post-processes Word files to fix right-to-left paragraph and run properties. |
| `raster_pdf.py` / `render_word.ps1` | Render a Word file to PDF/images for visual checks (the PowerShell script uses Word on Windows). |
| `proposal_hamrah_entekhab*.docx` | Generated proposal versions: base, 25-page expanded, final (B Nazanin) and Word RTL. |

## Requirements

Python 3.9+, `python-docx`, `Pillow`, `arabic-reshaper`, `python-bidi`. Paths in `build_proposal.py` point to the original Windows machine and must be adjusted.
