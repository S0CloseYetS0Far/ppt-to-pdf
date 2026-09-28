# PPT to PDF

Batch-convert PowerPoint files (`.ppt`, `.pptx`, `.pptm`, `.pps`, `.ppsx`) to PDF on Windows using Microsoft PowerPoint.

## Requirements
- Windows with Microsoft PowerPoint installed

## Usage
1. Put your PowerPoint files in the `input` folder.
2. Double-click `Convert.bat`.
3. PDFs appear in the `output` folder.

You can also drag files or folders onto `Convert.bat`.

### PowerShell options
```powershell
.\Convert-PptToPdf.ps1 -InputPath "C:\MyDecks" -Recurse -OutputDir "C:\PDFs" -Overwrite
```
- `-Recurse` – include subfolders
- `-Overwrite` – replace existing PDFs (skipped by default)

Originals are opened read-only and never modified.
