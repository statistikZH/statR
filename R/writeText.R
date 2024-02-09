#' Compact function to write and format text
#'
#' Facilitates the process of inserting text content into workbooks. This
#' function integrates operations openxlsx::writeData and openxlsx::addStyle,
#' providing a convenience function for text blocks.
#'
#' @description A shorthand function for
#' @param wb A workbook object
#' @param sheetname A sheetname
#' @param x Input
#' @param row The row at which to insert text
#' @param column The column at which to insert text. Can be a vector.
#' @param style A style object as defined by openxlsx::createStyle
#' @param name A name to assign to the text. Note that this has to be unique
#' @keywords writeText
#' @export
writeText <- function(wb, sheetname, x, row, column, style, name) {

  openxlsx::writeData(wb, sheetname, as.character(x), min(column), min(row),
    name = paste(sheetname, name, sep = "_"))

  row_extent <- namedRegionRowExtent(wb, sheetname, name)
  col_extent <- min(column):max(column)

  if (length(col_extent) > 1) {
    purrr::walk(row_extent,
                ~openxlsx::mergeCells(wb, sheetname, cols = col_extent, rows = .))
  }

  if (inherits(style, "Style")) {
    openxlsx::addStyle(
      wb, sheetname, style, row_extent, min(column))
  }

  if (!is.null(attr(x, "row_heights")) && length(attr(x, "row_heights")) == length(x)) {
    openxlsx::setRowHeights(wb, sheetname, row_extent, attr(x, "row_heights"))
  }
}
