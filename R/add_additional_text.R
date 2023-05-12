#' add_additional_text()
#'
#' @description Insert text into an existing workbook object. Text is inserted
#'  at column index 1.
#'
#' @returns Integer corresponding to the final row index
#'
#' @param wb Workbook object where text should be inserted
#'
#' @param text Character vector of text to be inserted
#'
#' @param sheetname Name of the worksheet where text should be inserted.
#'
#' @param start_row Row from which to insert elements
#'
#' @keywords internal
add_additional_text <- function(wb, text, sheetname, start_row){

  rows <- c(seq(start_row, start_row + length(text) - 1))

  # Write data
  openxlsx::writeData(wb, sheet = sheetname, x = text, colNames = FALSE,
                      headerStyle = style_subtitle(), startRow = start_row)

  # Apply style
  openxlsx::addStyle(wb, sheet = sheetname, style = style_subtitle(),
                     rows = rows, cols = 1, gridExpand = TRUE)

  return(max(rows))
}
