# Word Document Search Tool

Drop your Word (.docx) files and search by any keyword to get all matching content from every file.

## Setup

1. Open a terminal in this folder (`word-search-tool`).
2. Create a virtual environment (optional but recommended):
   ```bash
   python -m venv venv
   venv\Scripts\activate
   ```
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Run the tool

```bash
streamlit run app.py
```

Your browser will open with the app. You can:
- **Upload** one or more `.docx` files (drag & drop or click to browse).
- **Type a keyword** in the search box.
- **See all paragraphs** from any file that contain that keyword, with the filename shown for each result.

## Notes

- Only `.docx` (Word 2007+) format is supported.
- Text is taken from paragraphs and table cells.
- Search is case-insensitive.
