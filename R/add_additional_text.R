#' add_additional_text()
#'
#' @description Insert text into an existing workbook object.
#' @note: Text is inserted at first column index.
#' @param wb workbook object where text should be inserted
#' @param text Character vector of text to be inserted
#' @param sheetname Sheet where text should be inserted.
#' @param start_row Row from which to insert elements
#' @keywords internal

add_additional_text <- function(
  wb,
  text,
  sheetname,
  start_row
){

  rows = c(seq(start_row, start_row+length(text)-1))

  openxlsx::addStyle(wb,
                     sheet = sheetname,
                     style = style_subtitle(),
                     rows = rows,
                     cols = 1,
                     gridExpand = TRUE)

  openxlsx::writeData(wb,
                      sheet = sheetname,
                      x = text,
                      colNames = F,
                      headerStyle=style_subtitle(),
                      startRow = start_row)

  return(max(rows))
}
