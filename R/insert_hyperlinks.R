#' Insert hyperlinks with titles and sheetnames
#'
#' @description Function for inserting hyperlinks within an openxlsx Workbook
#' @inheritParams insert_worksheet
#' @param index_sheet_name Name of sheet where hyperlink should be created
#' @param sheetname Names of sheets to create hyperlinks to
#' @param title Titles of Hyperlinks
#' @param sheet_start_row Initial row after which hyperlinks should be created
#' @keywords insert_hyperlinks
#' @importFrom openxlsx makeHyperlinkString writeFormula addStyle
#' @export
insert_index_hyperlinks <- function(wb, sheetname, title,
                              index_sheet_name = "Index",
                              sheet_start_row = 15) {

  insert_hyperlinks(
    wb, sheetname, sheetname, index_sheet_name, sheet_start_row)

  openxlsx::writeData(
    wb, index_sheet_name, unlist(title), startCol = 7, startRow = sheet_start_row)

  purrr::walk(
    sheet_start_row + seq_along(sheetname) - 1,
    ~openxlsx::mergeCells(wb, index_sheet_name, 3:6, .x))
  purrr::walk(
    sheet_start_row + seq_along(sheetname) - 1,
    ~openxlsx::mergeCells(wb, index_sheet_name, 7:20, .x))
}


#' Insert hyperlinks
#'
#' @description Function for inserting hyperlinks within an openxlsx Workbook.
#'              Provides support for links to external .xlsx.
#' @inheritParams insert_worksheet
#' @param sheetname Names of sheets to create hyperlinks to
#' @param text Text to display in the cell with the hyperlink
#' @param where Name of the worksheet where the hyperlinks should be inserted
#' @param start_row First row from which hyperlink should be inserted
#' @param file External file. Default: NULL
#' @keywords insert_hyperlinks
#' @importFrom openxlsx makeHyperlinkString writeFormula addStyle
#' @export
insert_hyperlinks <- function(wb, sheetname, text, where,
                              start_row = 15, file = NULL) {

  # If File not found or not an xlsx, set to NULL and raise warning
  if (!is.null(file) && !(file.exists(file) || !grepl(".xlsx", file))) {
    warning("File not found or not an xlsx.")
  }

  hyperlink_strings <- makeHyperlinkString(sheetname, text = text,
                                           file = file)
  writeFormula(wb, where, hyperlink_strings, startCol = 3,
               startRow = start_row)
  addStyle(wb, where, hyperlinkStyle(),
           rows = start_row + seq_along(sheetname) - 1, cols = 3)

}
