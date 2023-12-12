#' insert_metadata_sheet()
#'
#' @description Function to add a formatted worksheet with metadata to an
#'   existing Workbook object.
#' @inheritParams insert_worksheet
#' @examples
#' # Create Workbook
#' wb <- openxlsx::createWorkbook()
#'
#' # Insert a simple metadata sheet
#' insert_metadata_sheet(wb, title = "Title of mtcars",
#'   metadata = c("Meta data information."))
#'
#' @keywords insert_metadata_sheet
#' @seealso createWorkbook, addWorksheet, writeData
#' @export
insert_metadata_sheet <- function(
    wb, sheetname = "Metadaten", title = "Title", source = getOption("statR_source"),
    metadata = NA, logo = getOption("statR_logo"),
    contactdetails = inputHelperContactInfo(compact = TRUE), author = "user") {

  # Add a new worksheet ------
  sheetname <- verifyInputSheetname(sheetname)
  openxlsx::addWorksheet(wb, sheetname)

  # Insert logo --------
  insert_worksheet_image(wb, sheetname, inputHelperLogoPath(logo), startrow = 1, startcol = 1)

  # Insert contact info, title, metadata, and sources into worksheet --------
  ### Contact info
  writeText(wb, sheetname, contactdetails, 2, 12, NULL, "contact")

  ### Request information
  writeText(wb, sheetname, paste(inputHelperDateCreated(), inputHelperAuthorName(author)),
            namedRegionLastRow(wb, sheetname, "contact") + 1, 12, NULL, "info")

  ### Headerline
  openxlsx::addStyle(wb, sheetname, style_headerline(), namedRegionLastRow(wb, sheetname, "contact") + 1,
                     1:26, gridExpand = TRUE, stack = TRUE)

  start_row <- namedRegionLastRow(wb, sheetname, "info") + 3

  if (is.character(title)){
    writeText(wb, sheetname, title, start_row, 1, style_title(), "title")
    start_row <- namedRegionLastRow(wb, sheetname, "title") + 1
  }

  if (is.character(source)) {
    writeText(wb, sheetname, source, start_row, 1, style_subtitle(), "source")
    start_row <- namedRegionLastRow(wb, sheetname, "source") + 1
  }

  if (is.character(metadata)) {
    writeText(wb, sheetname, metadata, start_row, 1, style_subtitle(), "metadata")
    start_row <- namedRegionLastRow(wb, sheetname, "metadata") + 1
  }

  ### Merge cells with title, metadata, and sources to ensure that they're displayed properly
  row_extent <- namedRegionRowExtent(wb, sheetname, c("title", "source", "metadata"))
  purrr::walk(row_extent, ~openxlsx::mergeCells(wb, sheetname, cols = 1:18, rows = .))

  ### Add Line wrapping
  openxlsx::addStyle(wb, sheetname, style_wrap(), row_extent, 1, stack = TRUE, gridExpand = TRUE)

  ### Hide gridlines
  openxlsx::showGridLines(wb, sheetname, FALSE)
}
