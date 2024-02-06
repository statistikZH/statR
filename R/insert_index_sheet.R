#' Insert empty index sheet
#'
#' @description Function which initializes an empty index sheet. The function is
#'   intended to be called at the beginning of a workflow (just after the
#'   initialization of the Workbook object), and paired with the function
#'   \code{insert_index_hyperlinks()}.
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
    openinghours = getOption("statR_openinghours"), source = getOption("statR_index_source"),
    author = "user") {

  insert_header(wb, sheetname, logo, contactdetails, homepage, auftrag_id, author,
                openinghours, contact_col = 15)

  ### Title - needs to exist
  writeText(wb, sheetname, title,
            namedRegionLastRow(wb, sheetname, "header_body") + 3,
            3:18, style_maintitle(), "title")

  ### Source - needs to exist
  writeText(wb, sheetname, source,
            namedRegionLastRow(wb, sheetname, "title") + 1,
            3:18, style_subtitle(), "source")

  ### Table of content caption
  writeText(wb, sheetname, getOption("statR_index_title"),
            namedRegionLastRow(wb, sheetname, "source") + 3,
            3, style_indextitle(), "toc")

  # Set column width of first column to 1 and hide gridlines
  openxlsx::setColWidths(wb, sheetname, 1, 1)
  openxlsx::showGridLines(wb, sheetname, FALSE)
}
