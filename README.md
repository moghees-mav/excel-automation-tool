# excel-automation-tool
it is a proof of skill project
A Python tool that cleans raw Excel data and generates automated summary reports for small business workflows
this tool will
- reads excel file
- cleans and formats data
- removes duplicates and empty rows
- generates summary

designed as a quick tool for business/indv who works with messy files

## How to use

1. Install requirements:
pip install -r requirements.txt

2. Place your Excel file anywhere (example: data/input.xlsx)

3. Run:
python src/main.py data/input.xlsx

4. Output file will be created at:
data/output.xlsx

With two sheets:
- CleanedData
- Summary

## Changelog

### v0.2.0 - 2026-02-09
- Added automatic column width adjustment based on content
- Added automatic row height adjustment based on content
- Column width capped at 50 characters for readability
- Row height scales with multi-line cell content
- Dynamic input file referencing via command-line argument
- File existence error handling with user-friendly messages
- Default input path (data/input.xlsx) when no argument provided

### v0.1.0 - Initial Release
- Excel file reading and writing
- Empty row removal
- Duplicate row removal
- Column name standardization (lowercase with underscores)
- Summary sheet generation with row/column counts and numeric column sums
- Output to separate sheets: CleanedData and Summary