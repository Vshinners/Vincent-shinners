#!/usr/bin/env Rscript

# Reimplementation of the MACRO_BKLOG VBA macro in R.
# It reads an Excel workbook, applies the same column manipulations and
# creates the filtered sheets produced by the VBA macro.

if (!requireNamespace("openxlsx", quietly = TRUE)) {
  stop("Package 'openxlsx' is required. Install with install.packages('openxlsx').")
}

excel_letter_to_index <- function(col) {
  col <- toupper(col)
  vapply(
    strsplit(col, ""),
    function(chars) {
      sum((utf8ToInt(chars) - 64) * 26 ^ (rev(seq_along(chars) - 1)))
    },
    integer(1)
  )
}

rename_by_letter <- function(df, letter, new_name) {
  idx <- excel_letter_to_index(letter)
  if (idx > ncol(df)) {
    stop(sprintf("Column %s (index %d) is missing from the sheet.", letter, idx))
  }
  names(df)[idx] <- new_name
  df
}

delete_by_letters <- function(df, letters) {
  idx <- excel_letter_to_index(letters)
  idx <- sort(unique(idx))
  if (length(idx) == 0) return(df)
  if (max(idx) > ncol(df)) {
    stop("Attempting to delete a column that does not exist.")
  }
  df[, -idx, drop = FALSE]
}

move_column_by_letter <- function(df, from_letter, dest_letter) {
  from_idx <- excel_letter_to_index(from_letter)
  if (from_idx > ncol(df)) stop(sprintf("Column %s is missing.", from_letter))
  from_name <- names(df)[from_idx]
  df_wo <- df[, -from_idx, drop = FALSE]
  dest_idx <- excel_letter_to_index(dest_letter)
  dest_idx <- min(dest_idx, ncol(df_wo) + 1)
  new_order <- append(names(df_wo), from_name, after = dest_idx - 1)
  df_wo[, new_order, drop = FALSE]
}

move_range_by_letter <- function(df, start_letter, end_letter, dest_letter) {
  start_idx <- excel_letter_to_index(start_letter)
  end_idx <- excel_letter_to_index(end_letter)
  rng <- start_idx:end_idx
  if (max(rng) > ncol(df)) stop("Range to move exceeds available columns.")
  move_names <- names(df)[rng]
  df_wo <- df[, -rng, drop = FALSE]
  dest_idx <- excel_letter_to_index(dest_letter)
  dest_idx <- min(dest_idx, ncol(df_wo) + 1)
  new_order <- append(names(df_wo), move_names, after = dest_idx - 1)
  df_wo[, new_order, drop = FALSE]
}

process_backlog <- function(input_path, sheet_name = "Sheet1") {
  df <- openxlsx::read.xlsx(input_path, sheet = sheet_name, colNames = TRUE)
  df <- data.frame(df, check.names = FALSE, stringsAsFactors = FALSE)

  # Ensure there are enough columns for the operations below.
  if (ncol(df) < excel_letter_to_index("BM")) {
    stop("Sheet does not contain the expected number of columns for the macro.")
  }

  # Sort by columns O then T (as in the VBA macro).
  df <- df[order(df[[excel_letter_to_index("O")]], df[[excel_letter_to_index("T")]]), , drop = FALSE]

  # Header updates and early deletions.
  df <- rename_by_letter(df, "O", "PDD")
  df <- rename_by_letter(df, "C", "Org Qty")
  df <- rename_by_letter(df, "D", "Rem Qty")
  df <- rename_by_letter(df, "E", "NOTES")

  df <- delete_by_letters(df, "F")

  rename_map <- c("F" = "Ax12", "G" = "Ax09", "H" = "INV", "I" = "All WHS",
                  "J" = "AX12 WIP", "K" = "12 WIP", "L" = "09 WIP")
  for (letter in names(rename_map)) {
    df <- rename_by_letter(df, letter, rename_map[[letter]])
  }

  df <- delete_by_letters(df, "M")

  df <- rename_by_letter(df, "Q", "SHC")
  df <- move_column_by_letter(df, "Q", "B")

  df <- delete_by_letters(df, c("S", "T"))
  df <- rename_by_letter(df, "S", "BAAN")
  df <- delete_by_letters(df, "T")
  df <- rename_by_letter(df, "T", "Org $")
  df <- rename_by_letter(df, "U", "Rem $")
  df <- delete_by_letters(df, c("V", "W"))
  df <- rename_by_letter(df, "V", "Unit $")

  df <- rename_by_letter(df, "AD", "RATING")
  df <- rename_by_letter(df, "AE", "SLDC")
  df <- rename_by_letter(df, "AF", "D/C rest")
  df <- rename_by_letter(df, "AG", "LOT INFO")
  df <- rename_by_letter(df, "AR", "DPA")

  df <- move_column_by_letter(df, "AR", "B")
  df <- move_column_by_letter(df, "Z", "C")
  df <- move_column_by_letter(df, "AX", "AH")
  df <- move_range_by_letter(df, "BL", "BM", "AB")
  df <- move_range_by_letter(df, "AA", "AK", "V")

  df
}

write_backlog_output <- function(df, output_path) {
  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, "Sheet1")
  openxlsx::writeData(wb, "Sheet1", df)

  add_filtered <- function(name, values) {
    openxlsx::addWorksheet(wb, name)
    openxlsx::writeData(
      wb,
      name,
      df[df[[1]] %in% values, , drop = FALSE]
    )
  }

  add_filtered("ALL JANS", c("PE18", "PE30", "PE34", "PE39"))
  add_filtered("METAL", c("PE22", "PE23", "PE29", "PE31", "PE32", "PE35", "PE36", "PE38"))
  add_filtered("Modules", "PE21")
  add_filtered("RECTIFIERS", c("PE06A", "PE06C", "PE06AS", "PE06ES"))

  openxlsx::saveWorkbook(wb, output_path, overwrite = TRUE)
}

main <- function() {
  args <- commandArgs(trailingOnly = TRUE)
  input_path <- if (length(args) >= 1) args[1] else "Metal Total Forcast.xlsx"
  output_path <- if (length(args) >= 2) args[2] else "MACRO_BKLOG_output.xlsx"
  sheet_name <- if (length(args) >= 3) args[3] else "Sheet1"

  df <- process_backlog(input_path, sheet_name)
  write_backlog_output(df, output_path)
}

if (!interactive()) {
  main()
}
