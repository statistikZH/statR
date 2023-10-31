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
  insert_worksheet_image(
    wb = wb, sheetname = sheetname, image = inputHelperLogoPath(logo),
    startrow = 1, startcol = 1)

  # Insert contact info, title, metadata, and sources into worksheet --------
  ### Contact info
  openxlsx::writeData(
    wb, sheetname, contactdetails, 12, 2, name = paste(sheetname, "contact", sep = "_"))

  ### Request information
  openxlsx::writeData(
    wb, sheetname, x = paste(inputHelperDateCreated(), inputHelperAuthorName(author)),
    startCol = 12, startRow = namedRegionLastRow(wb, sheetname, "contact") + 1,
    name = paste(sheetname, "info", sep = "_"))

  ### Headerline
  openxlsx::addStyle(
    wb, sheetname, style_headerline(),
    namedRegionLastRow(wb, sheetname, "contact") + 1, 1:26,
    gridExpand = TRUE, stack = TRUE)

  ### Title
  openxlsx::writeData(
    wb, sheetname, title, startRow = namedRegionLastRow(wb, sheetname, "info") + 3,
    name = paste(sheetname, "title", sep = "_"))

  openxlsx::addStyle(
    wb, sheetname, style_title(), rows = namedRegionLastRow(wb, sheetname, "title"),
    cols = 1)

  if (!is.null(source) && !all(is.na(source))) {
    openxlsx::writeData(
      wb, sheetname, inputHelperSource(source, collapse = NULL),
      startRow = namedRegionLastRow(wb, sheetname, "title") + 1,
      name = paste(sheetname, "source", sep = "_"))
  }

  if (!is.null(metadata) && !all(is.na(metadata))) {
    openxlsx::writeData(
      wb, sheetname, inputHelperMetadata(metadata, collapse = NULL),
      startRow = namedRegionLastRow(wb, sheetname, "source") + 1,
      name = paste(sheetname, "metadata", sep = "_"))
  }

  ### Merge cells with title, metadata, and sources to ensure that they're displayed properly
  purrr::walk(namedRegionRowExtent(wb, sheetname, c("title", "source", "metadata")),
              ~openxlsx::mergeCells(wb, sheetname, cols = 1:18, rows = .))

  ### Add Line wrapping
  openxlsx::addStyle(wb, sheetname, style_wrap(),
                     namedRegionRowExtent(wb, sheetname, c("title", "source", "metadata")), 1,
                     stack = TRUE, gridExpand = TRUE)


  # openxlsx::addStyle(wb, sheetname, style_subtitle(),
  #                    c(namedRegionFirstRow(wb, sheetname, "source"),
  #                      namedRegionFirstRow(wb, sheetname, "metadata")), 1)



  ### Hide gridlines
  openxlsx::showGridLines(wb, sheetname, FALSE)
}
