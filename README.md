# Magic 2026 Web Converter

## Behavior

- User uploads workbook (`.xlsm`) every run.
- User uploads SysCAD text file (`.txt`) every run.
- Template is loaded by default from `default_template.json` (no required template upload).
- Optional template override upload is available.

## Run Locally

Use any static web server from the `webapp` folder.

Example with Node:

```bash
npx serve .
```

Then open the local URL shown in terminal.

## Deploy To GitHub Pages

1. Create a GitHub repository and push the `webapp` folder contents to the root (or `/docs`).
2. In repository settings, open Pages.
3. Set source branch/folder to where `index.html` is located.
4. Save and wait for deployment.
5. Open the GitHub Pages URL.

## Files

- `index.html` UI
- `styles.css` styling
- `app.js` conversion logic in browser
- `default_template.json` default template

## Notes

- No Python runtime is needed for end users.
- No `upd.exe` is used.
- Conversion runs in the browser, files stay local to user session.
