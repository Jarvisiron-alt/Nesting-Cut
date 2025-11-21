# Nesting App - PDF Report Export

This project adds a formatted PDF export (similar to the shared sample) to the Streamlit-based nesting tool.

## New: Report-Style PDF
- Header bar with logo (defaults to a simple "Ei8" logo if none provided).
- Title (configurable).
- Job metadata table (material, thickness, sheet size, kerf, spacing, rotation step, total sheets).
- Per-sheet section with utilization, parts count, preview image, and parts table.

## Usage
1. Install dependencies:
```pwsh
pip install -r requirements.txt
```
2. Run the app:
```pwsh
streamlit run app.py
```
3. In the CUT stage -> Export -> choose `PDF (.pdf)`.
4. Enable "Use report layout", optionally upload your logo (PNG/JPG), set the title, choose scope (single or all sheets), then click `Prepare PDF`.

If ReportLab is not installed, the app falls back to the previous image-only PDF export.

## Notes
- The logo area expects a square-ish image. If no logo is uploaded, a placeholder "Ei8" logo is rendered.
- The per-sheet preview uses the same internal preview generator as the UI.
- If you want the report to exactly match a specific layout, share exact fields and desired arrangement; we can adjust the template accordingly.
