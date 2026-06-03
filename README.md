# Malbon China Y26–Y27 Proposal

An AI-generated HTML recreation of the Malbon × WalktheChat "Malbon China Y26–Y27 Proposal" deck, plus the structured data extracted from the original PowerPoint that it was built from.

## Contents

- **`malbon-ai-generated-proposal/`** — the AI-generated deck.
  - `Malbon China Proposal.html` — interactive deck (uses `deck-stage.js`).
  - `Malbon China Proposal-print.html` — flat, print-friendly version.
  - `assets/` — per-slide imagery (`slideNN/`).
  - `review/`, `scratchpad.md` — build notes and review renders.
- **`final_json/`** — verbatim structured content of the original 49 slides (`slides_01-08.json` … `slides_41-49.json`), including titles, chart values, callouts, image descriptions, and per-slide logical intent.
- **`images/`** — images extracted per slide from the source deck.
- **`malbon_proposal_structured.json`**, **`malbon_raw_extract.json`** — intermediate extraction outputs.
- **`extract.py`** — script used to extract content from the source `.pptx`.

## Notes

The 49-slide AI deck maps 1:1 to the source deck. The source `.pptx` and its unzipped `extracted/` folder are intentionally not committed (see `.gitignore`).

### Slide 3 & Slide 6 chart corrections

- **Slide 3** — both market charts were rebuilt as zero-based SVG charts so the visual contrast (steady on-course growth vs. fast indoor growth) honestly matches the +5.1% vs +118% figures.
- **Slide 6** — the social-assessment chart was corrected to the real source values (RED mention volume + YoY growth per brand) after the initial AI version contained fabricated figures.
