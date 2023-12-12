

#' Compact function to write text into a worksheet and apply a style
#'
#' @description A shorthand function for openxlsx::writeData and openxlsx::addStyle
#' @param wb A workbook object
#' @param sheetname A sheetname
#' @param x Input
#' @param row The row at which to insert text
#' @param column The column at which to insert text
#' @param style A style object as defined by openxlsx::createStyle
#' @param name A name to assign to the text. Note that this has to be unique
#' @keywords writeText
#' @export
writeText <- function(wb, sheetname, x, row, column, style, name) {
  openxlsx::writeData(
    wb, sheetname, as.character(x), column, row,
    name = paste(sheetname, name, sep = "_"))

  if (inherits(style, "Style")) {
    openxlsx::addStyle(
      wb, sheetname, style, namedRegionRowExtent(wb, sheetname, name), column)
  }

  if (!is.null(attr(x, "row_heights")) && length(attr(x, "row_heights")) == length(x))
    openxlsx::setRowHeights(wb, sheetname, row + seq_along(x) - 1, attr(x, "row_heights"))
}

