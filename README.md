
# Flowcell QC

A Streamlit app that reads flowcell HTML reports, previews Lane QC and
Barcode QC tables, and downloads one Excel workbook with two tabs.

## Run on Windows

Open a terminal in this folder and run:

```powershell
py -3.14 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m streamlit run app.py
```
