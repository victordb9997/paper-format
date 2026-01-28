# Nature Format Studio

Nature Format Studio is a lightweight Flask app that turns a DOCX manuscript and a PPTX figure deck into a Nature-inspired web layout and PDF.

## Quick start

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app/app.py
```

Then visit http://localhost:5000, upload your manuscript and figure deck, and download the formatted PDF.

## Notes

- The manuscript parser uses the first paragraph as the title and the second paragraph as the abstract.
- Figures are pulled from images embedded in the PowerPoint slides.
- The PDF export uses ReportLab and does not require system GTK dependencies.
- Customize the layout by editing `app/static/styles.css` and `app/templates/preview.html`.

## Troubleshooting (Windows)

If you previously installed WeasyPrint and see errors about missing `gobject-2.0-0` or GTK, delete the old virtual environment and reinstall dependencies:

```bash
deactivate
rm -rf .venv
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```
