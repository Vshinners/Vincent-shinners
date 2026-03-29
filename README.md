# Vincent-shinners

## MACRO_BKLOG VBA conversion

The VBA macro from `MACRO_BKLOG.docx` has been reimplemented in `macro_bklog.R`.
It reads an Excel workbook, performs the same column reordering/renaming, and
creates the filtered sheets produced by the macro.

Usage:

```sh
Rscript macro_bklog.R "<input workbook.xlsx>" "<output workbook.xlsx>" "Sheet1"
```

Defaults:
- input workbook: `Metal Total Forcast.xlsx`
- output workbook: `MACRO_BKLOG_output.xlsx`
- sheet name: `Sheet1`

Requires the `openxlsx` R package (`install.packages("openxlsx")` if missing).
