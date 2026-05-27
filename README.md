# Nutrition Label Extraction Pipeline

Extracts **Nutrition Facts** panels and **ingredient statements** from product
marketing / specification PDFs and images of nutrition labels and writes them to an Excel workbook for
integration with **FPED** and **FNDDS**.

## What it does

1. Reads every PDF and image file in `pdfs/` (or input folder specified with `-i <foldername>/`)
2. Recovers the Nutrition Facts panel:
   * **Vector-text PDFs** - PDF reading order scrambles the panel, so visual
     rows are rebuilt from word coordinates and each label is kept next to its
     value.
   * **Image / broken-font panels** - rendered and read with Tesseract OCR.
3. Parses macro- and micronutrients, the serving size, and the ingredient list.
4. Back-calculates micronutrient amounts that a label only prints as a
   **% Daily Value** (common for B-vitamins on cereals) using the FDA Daily
   Value table. These are flagged in `validation_notes`.
5. Converts every nutrient to a **per-100 g** basis (the basis FNDDS uses).
6. Runs validation checks (Atwater energy cross-check, added vs total sugars,
   implausible per-100 g values, OCR-sourced rows) and records them.

## Output

`output/nutrition_extraction.xlsx` (or output path specified with `-o <foldername>/<filename>.xlsx`), with two sheets:

* **Extraction** - one row per PDF: source / product metadata, serving size,
  per-serving amounts, per-100 g amounts, `validation_notes`, `ingredients`.
* **FNDDS_per_100g** - the per-100 g values laid out with FNDDS nutrient
  numbers and descriptions in the column headers.

Every row also carries a `confidence` (0-100) and `confidence_rating`
(`High` >=85, `Medium` 65-84, `Low` 1-64, `Failed` 0) so quality checkers can
sort the worksheet and review the lowest-confidence rows first. Penalties
stack from the same signals the validator records (calories or serving size
missing, Atwater mismatch, OCR usage, %DV back-fill, OCR rescaling,
implausible per-100 g values, missing ingredients).

## Requirements

```
pip install -r requirements.txt
```

Plus **Tesseract OCR** (used only for the image-based panels). No PATH change
is needed - the script finds Tesseract via the `TESSERACT_CMD` environment
variable, then the system PATH, then the default Windows install location.

## Run

```
python nutrition_extract.py
```

## Notes / limitations

* Standard FDA panels list ~15 nutrients; cereals list ~25. Nutrients not
  printed on a panel are left blank.
* Rows extracted by OCR are flagged `panel read by OCR - verify against source
  PDF`; OCR can misread small digits, so spot-check those products.
* `validation_notes` is the audit trail - always review it before using a row.

## TODO

* Direct FNDDS load file (food codes, FNDDS nutrient-file layout)
