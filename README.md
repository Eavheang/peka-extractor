# peka-extractor

Extracts PDF, DOCX, XLSX, PPTX, TXT, and Markdown files into JSON.

## Run

Requires Python 3.11 or newer.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python extractor.py .\test_data\sample.pdf
```

Output is written to `.\extracted\`. Pass multiple files or choose another output directory:

```powershell
python extractor.py .\file.pdf .\notes.docx --output .\output
```

Run `python extractor.py --help` for all options.
