#' add_additional_text()
#'
#' Insert text into an existing workbook object. Note that text is inserted
#' at first column index.
#' @param wb workbook object where text should be inserted
#' @param text Character vector of text to be inserted
#' @param sheetname Sheet where text should be inserted.
#' @param start_row Row from which to insert elements
#' @keywords add_additional_text
#' @examples
#'
#' # Generation of a spreadsheet
#' wb <- openxlsx::createWorkbook("hello")
#'
#' insert_metadata_sheet(wb, title = "Title of mtcars", metadata = c("Meta data information."))


add_additional_text <- function(wb, text, sheetname, start_row){

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
