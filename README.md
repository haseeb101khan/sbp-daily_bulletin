# Daily Bulletin System - Vercel Preview

This is the Vercel-ready preview version of the Excel-to-Word Daily Bulletin MVP.

## Workflow

1. Open the deployed Vercel site.
2. Download the empty Excel template.
3. Fill the `News Input` sheet.
4. Upload the completed `.xlsx`; the upload box turns green when the file is ready.
5. Click **Generate Report**.
6. Use **View Report** to review the bulletin in the browser.
7. Download the generated `.docx` when the preview looks correct.

## Deploy To Vercel

Install Vercel CLI, then run these commands from this folder:

```powershell
vercel dev
```

For deployment:

```powershell
vercel
```

## How This Version Works

The static frontend is served from `index.html` and `static/`. The Python serverless function at `api/generate.py` receives the Excel upload, generates a preview HTML file and Word document in temporary storage, then returns both to the browser in one response.

`pyproject.toml` points Vercel at the Python entrypoint:

```toml
[tool.vercel]
entrypoint = "api.generate:handler"
```

This version is meant for previews and initial testing. For production, use persistent storage such as Vercel Blob, an internal server, or a packaged local Windows app.
