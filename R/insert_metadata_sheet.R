#' Insert sheet with metadata
#'
#' @description Function to add a formatted worksheet with metadata to an
#'   existing Workbook object.
#' @inheritParams insert_worksheet
#' @param meta_infos a list with title, source, and metadata as named objects
#' @keywords insert_metadata_sheet
#' @seealso createWorkbook, addWorksheet, writeData
#' @export
insert_metadata_sheet <- function(
    wb, sheetname, meta_infos, logo = getOption("statR_logo"),
    contactdetails = inputHelperContactInfo(compact = TRUE), author = "user") {

  title <- meta_infos[["title"]]
  source <- meta_infos[["source"]]
  metadata <- meta_infos[["metadata"]]

  insert_header(wb, sheetname, logo, contactdetails, NULL, NULL, author, NULL, 15)

  start_row <- namedRegionLastRow(wb, sheetname, "info") + 3


  if (is.character(title)){
    writeText(wb, sheetname, title, start_row, 1:18, style_title(), "title")
    start_row <- namedRegionLastRow(wb, sheetname, "title") + 1
  }

  if (is.character(source)) {
    writeText(wb, sheetname, source, start_row, 1:18, style_subtitle(), "source")
    start_row <- namedRegionLastRow(wb, sheetname, "source") + 1
  }

  if (is.character(metadata)) {
    writeText(wb, sheetname, metadata, start_row, 1:18, style_subtitle(), "metadata")
    start_row <- namedRegionLastRow(wb, sheetname, "metadata") + 1
  }

  openxlsx::showGridLines(wb, sheetname, FALSE)
}
