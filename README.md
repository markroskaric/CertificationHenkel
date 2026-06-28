# CertificationHenkel

## What this program does

`certificateFinder.ps1` scans a folder of PDF certificates and compares the filenames against the Excel table in `BatchIdh.xlsx`.
It generates a separate Excel report with the matching results and status colors.

## Files and locations

- `certificateFinder.ps1`
  - The main PowerShell UI script.
  - It must stay in the same folder as `BatchIdh.xlsx`.
- `BatchIdh.xlsx`
  - The input workbook containing the table of `IDH` and `Batch` values.
  - Must be located in the same directory as `certificateFinder.ps1`.
- `Certificates/`
  - Example folder for your PDF certificate files.
  - You should browse to the folder containing your PDF files when running the app.
- `CertificateReport_*.xlsx`
  - The generated report file.
  - It is created in the same directory as `certificateFinder.ps1`.

## How to run

### Option 1: Run from File Explorer (recommended)

1. Open the folder containing `certificateFinder.ps1`.
2. Right-click `certificateFinder.ps1`.
3. Choose `Run with PowerShell`.
4. The UI window will open.
5. In the UI:
   - Click `Browse` and choose the folder containing your PDF certificate files.
   - Click `START`.
6. When finished, the script will create a report file named `CertificateReport_YYYYMMDD_HHMMSS.xlsx` in the same folder.

### Option 2: Run from PowerShell prompt

1. Open PowerShell.
2. Change to this repository folder:
   ```powershell
   cd "c:\Users\yourUser\Documents\GitHub\CertificationHenkel"
   ```
3. Run the script:
   ```powershell
   .\certificateFinder.ps1
   ```
4. In the UI:
   - Click `Browse` and choose the folder containing your PDF certificate files.
   - Click `START`.
5. When finished, the script will create a report file named `CertificateReport_YYYYMMDD_HHMMSS.xlsx`.

## Expected Excel input format

The workbook `BatchIdh.xlsx` must contain a table with at least these columns:

- `IDH`
- `Batch`

The script reads the table rows and compares each `IDH|Batch` combination against PDF filenames.

## Status codes in the report

The generated report uses these statuses:

- `Non-08`
  - The `Batch` value does not start with `08`.
  - This row is ignored for PDF matching.
- `Duplicate`
  - The same `IDH|Batch` combination appears more than once in the table.
  - Only the second and later occurrences are marked as duplicates.
- `Found`
  - A matching PDF was located for that `IDH|Batch`.
- `Missing`
  - No matching PDF was found for that `IDH|Batch`.

## Notes

- The script does not modify `BatchIdh.xlsx`.
- The report file is saved in the same folder as `certificateFinder.ps1`.
- PDF paths in the report are clickable when opened in Excel.

