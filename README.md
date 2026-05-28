# SAPUI5 Excel Demo

Small demo app with one screen and three buttons:

- Execute — shows a message and updates the result text.
- Template Excel — downloads a simple Excel template (`template.xlsx`).
- Upload Excel — opens file picker; when an Excel file is selected it is parsed and displayed.

Run locally:

Open `index.html` in your browser (double-click or use `start` on Windows):

```powershell
start sap-excel-app\index.html
```

Or run a simple HTTP server from the workspace root and open http://localhost:8000/sap-excel-app/index.html:

```powershell
python -m http.server 8000
```
