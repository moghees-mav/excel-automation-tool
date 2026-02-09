# excel-automation-tool
it is a proof of skill project
A Python tool that cleans raw Excel data and generates automated summary reports for small business workflows
this tool will
- reads excel file
- cleans and formats data
- removes duplicates and empty rows
- generates summary

designed as a quick tool for business/indv who works with messy files

##status
under development

## Changelog

### v0.2.0 - 2026-02-09
- Added automatic column width adjustment based on content
- Added automatic row height adjustment based on content
- Column width capped at 50 characters for readability
- Row height scales with multi-line cell content

### v0.1.0 - Initial Release
- Excel file reading and writing
- Empty row removal
- Duplicate row removal
- Column name standardization (lowercase with underscores)
- Summary sheet generation with row/column counts and numeric column sums
- Output to separate sheets: CleanedData and Summary