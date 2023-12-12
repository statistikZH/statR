#' insert_index_sheet()
#'
#' @description Function which generates an index sheet inside an openxlsx
#'   workbook.
#' @details The logo can be customized via the `statR_logo` global option. This
#'   should point to the path of the logo file. Other options such as the image
#'   size in the final .xlsx can either be changed via the options
#'   `statR_logo_width` and `statR_logo_height`, or set along with contact
#'   information in a custom profile.
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


  # # Insert logo ----------
  insert_worksheet_image(wb, sheetname, inputHelperLogoPath(logo), startrow = 1, startcol = 1)

  # Insert contact info, title, metadata, and sources into worksheet --------
  ### Contact information
  writeText(wb, sheetname, contactdetails, 2, 15, NULL, "contact")

  ### Office hours
  writeText(wb, sheetname, openinghours, namedRegionFirstRow(wb, sheetname, "contact"),
            18, NULL, "officehours")

  # ### Homepage
  writeText(wb, sheetname, inputHelperHomepage(homepage), namedRegionLastRow(wb, sheetname, "contact") + 1,
            15, NULL, "homepage")

  ### Request information
  writeText(wb, sheetname, c(inputHelperDateCreated(), inputHelperOrderNumber(auftrag_id)),
            namedRegionLastRow(wb, sheetname, "homepage") + 3, 15, NULL, "info")

  ### Add Headerline
  openxlsx::addStyle(wb, sheetname, style_headerline(),
    namedRegionLastRow(wb, sheetname, "info") + 1, 1:20,
    gridExpand = TRUE, stack = TRUE)


  ### Title - needs to exist
  writeText(wb, sheetname, title, namedRegionLastRow(wb, sheetname, "info") + 3,
            3, style_maintitle(), "title")

  ### Source - needs to exist
  writeText(wb, sheetname, source, namedRegionLastRow(wb, sheetname, "title") + 1,
            3, style_subtitle(), "source")

  # Merge cells
  row_extent <- namedRegionRowExtent(wb, sheetname, c("title", "source"))
  purrr::walk(row_extent, ~openxlsx::mergeCells(wb, sheetname, cols = 3:18, rows = .))


  ### Table of content caption
  writeText(wb, sheetname, getOption("statR_toc_title"), namedRegionLastRow(wb, sheetname, "source") + 3,
            3, style_indextitle(), "toc")

  # Format ---------
  ### Set column width of first column to 1
  openxlsx::setColWidths(wb, sheetname, 1, 1)

  ### Hide gridlines
  openxlsx::showGridLines(wb, sheetname, FALSE)
}
