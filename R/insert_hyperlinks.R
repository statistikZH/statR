#' insert_index_hyperlinks()
#'
#' @description Function for inserting hyperlinks within an openxlsx Workbook
#' @inheritParams insert_worksheet
#' @param sheet_row Row where hyperlink should be inserted
#' @param index_sheet_name Name of sheet where hyperlink should be created
#' @param sheetnames Names of sheets to create hyperlinks to
#' @param titles Titles of Hyperlinks
#' @param sheet_start_row Initial row after which hyperlinks should be created
#' @keywords insert_hyperlinks
#' @importFrom openxlsx makeHyperlinkString writeFormula addStyle
#' @export
insert_index_hyperlinks <- function(wb, sheetnames, titles,
                              index_sheet_name = "Index",
                              sheet_start_row = 15) {

  insert_hyperlinks(wb, sheetnames, titles, index_sheet_name, sheet_start_row,
                    NULL)
}


#' insert_hyperlinks()
#'
#' @description Function for inserting hyperlinks within an openxlsx Workbook.
#'              Provides support for links to external .xlsx
#' @inheritParams insert_worksheet
#' @param sheet_row Row where hyperlink should be inserted
#' @param index_sheet_name Name of sheet where hyperlink should be created
#' @param sheetnames Names of sheets to create hyperlinks to
#' @param titles Titles of Hyperlinks
#' @param sheet_start_row Initial row after which hyperlinks should be created
#' @keywords insert_hyperlinks
#' @importFrom openxlsx makeHyperlinkString writeFormula addStyle
#' @export
insert_hyperlinks <- function(wb, sheetnames, text, where,
                              start_row, file = NULL) {

  # If File not found or not an xlsx, set to NULL and raise warning
  if (!is.null(file) && !(file.exists(file) && grepl(".xlsx", file))) {
    warning("File not found or not an xlsx. Set to NULL")
    file <- NULL
  }

  hyperlink_strings <- makeHyperlinkString(sheetnames, text = text,
                                           file = file)
  writeFormula(wb, where, hyperlink_strings, startCol = 3,
               startRow = start_row)
  addStyle(wb, where, hyperlinkStyle(),
           rows = start_row + seq_along(sheetnames) - 1, cols = 3)

}

