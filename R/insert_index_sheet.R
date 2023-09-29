#' insert_index_sheet()
#'
#' @description Function which generates an index sheet inside an openxlsx
#'   workbook.
#' @inheritParams insert_worksheet
#' @param openinghours statzh or a character string or vector with opening hours
#' @param auftrag_id Order ID
#' @keywords insert_index_sheet
#' @export
insert_index_sheet <- function(
    wb, sheetname = "Index", title, auftrag_id, logo = getOption("statR_logo"),
    contactdetails = inputHelperContactInfo(), homepage = getOption("statR_homepage"),
    openinghours = getOption("statR_openinghours"), source = getOption("statR_source")) {

  # Initialize new worksheet as index sheet ------
  openxlsx::addWorksheet(wb, sheetname)


  # Insert logo ----------
  insert_worksheet_image(
    wb = wb, sheetname = sheetname, image = inputHelperLogoPath(logo), startrow = 1,
    startcol = 1, width = 2.145, height = 0.7865)


  # Insert contact info, title, metadata, and sources into worksheet --------
  ### Contact information
  openxlsx::writeData(
    wb, sheetname, contactdetails, startCol = 15, startRow = 2,
    name = paste(sheetname,"contact", sep = "_"))

  ### Office hours
  openxlsx::writeData(
    wb, sheetname, openinghours, startCol = 18,
    startRow = namedRegionFirstRow(wb, sheetname, "contact"),
    name = paste(sheetname, "officehours", sep = "_"))

  ### Homepage
  openxlsx::writeData(
    wb, sheetname, inputHelperHomepage(homepage), startCol = 15,
    startRow = namedRegionLastRow(wb, sheetname, "contact") + 1,
    name = paste(sheetname, "homepage", sep = "_"))


  ### Add Headerline
  openxlsx::addStyle(
    wb, sheetname, style_headerline(), namedRegionLastRow(wb, sheetname, "homepage") + 1, 1:20,
    gridExpand = TRUE, stack = TRUE)

  ### Request information
  openxlsx::writeData(
    wb, sheetname, c(inputHelperDateCreated(), inputHelperOrderNumber(auftrag_id)),
    startCol = 15, startRow = namedRegionLastRow(wb, sheetname, "homepage") + 3,
    name = paste(sheetname, "info", sep = "_"))

  ### Title
  openxlsx::writeData(
    wb, sheetname, title, 3, namedRegionLastRow(wb, sheetname, "info") + 1,
    name = paste(sheetname, "title", sep = "_"))

  openxlsx::addStyle(
    wb, sheetname, style_maintitle(), namedRegionLastRow(wb, sheetname, "title"), 3)

  ### Source
  openxlsx::writeData(
    wb, sheetname, inputHelperSource(source, collapse = "; "), startCol = 3,
    startRow = namedRegionLastRow(wb, sheetname, "title") + 1,
    name = paste(sheetname, "source", sep = "_"))

  ### Table of content caption
  openxlsx::writeData(
    wb, sheetname, getOption("statR_toc_title"), 3, namedRegionLastRow(wb, sheetname, "source") + 3,
    name = paste(sheetname, "toc", sep = "_"))

  openxlsx::addStyle(
    wb, sheetname, subtitleStyle(), namedRegionLastRow(wb, sheetname, "toc"), 3)

  purrr::walk(
    namedRegionFirstRow(wb, sheetname, "title"):namedRegionFirstRow(wb, sheetname, "toc"),
    ~openxlsx::mergeCells(wb, sheetname, cols = 3:13, rows = .x))

  purrr::walk(
    namedRegionFirstRow(wb, sheetname, "title"):namedRegionFirstRow(wb, sheetname, "toc"),
    ~openxlsx::addStyle(wb, sheetname, style_wrap(), cols = 3:13, rows = .x, stack = TRUE))

  # Format ---------
  ### Set column width of first column to 1
  openxlsx::setColWidths(wb, sheetname, 1, 1)

  ### Hide gridlines
  openxlsx::showGridLines(wb, sheetname, FALSE)
}
