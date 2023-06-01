#' insert_index_sheet()
#'
#' @description Function which generates an index sheet inside an openxlsx
#'   workbook.
#' @inheritParams insert_worksheet
#' @param openinghours statzh or a character string or vector with opening hours
#' @param auftrag_id Order ID
#' @param order_id_prefix Text inserted before order ID
#' @param toc_caption Text for table of content
#' @keywords insert_index_sheet
#' @export
insert_index_sheet <- function(wb,
                               sheetname = "Index",
                               title,
                               auftrag_id,
                               logo = "statzh",
                               contactdetails = "statzh",
                               homepage = "statzh",
                               openinghours = "statzh",
                               source = "statzh",
                               date_prefix = "Erstellt am: ",
                               order_id_prefix = "Auftragsnr.:",
                               toc_caption = "Inhalt"){

  # Initialize new worksheet as index sheet ------
  openxlsx::addWorksheet(wb, sheetname)


  # Insert logo ----------
  insert_worksheet_image(wb = wb, sheetname = sheetname,
                         image = inputHelperLogoPath(logo),
                         startrow = 1, startcol = 1,
                         width = 2.145, height = 0.7865)


  # Insert contact info, title, metadata, and sources into worksheet --------
  ### Contact information
  openxlsx::writeData(wb, sheetname,
                      x = c(inputHelperContactInfo(contactdetails),
                            inputHelperHomepage(homepage)),
                      startCol = 15, startRow = 2,
                      name = paste(sheetname,"contact", sep = "_"))

  ### Office hours
  openxlsx::writeData(wb, sheetname, inputHelperOfficeHours(openinghours),
                      18, namedRegionFirstRow(wb, sheetname, "contact"),
                      name = paste(sheetname,"officehours", sep = "_"))

  ### Request information
  openxlsx::writeData(wb, sheetname,
                      c(inputHelperDateCreated(prefix = date_prefix),
                        inputHelperOrderNumber(auftrag_id, prefix = order_id_prefix)),
                      startCol = 15, startRow = namedRegionLastRow(wb, sheetname, "contact"),
                      name = paste(sheetname,"info", sep = "_"))
  ### Add Headerline
  openxlsx::addStyle(wb, sheetname, style_headerline(),
                     namedRegionLastRow(wb, sheetname, "info"), 1:20,
                     gridExpand = TRUE, stack = TRUE)

  ### Title
  openxlsx::writeData(wb, sheetname, title, 3,
                      namedRegionLastRow(wb, sheetname, "info") + 3,
                      name = paste(sheetname,"title", sep = "_"))
  openxlsx::addStyle(wb, sheetname, style_maintitle(),
                     namedRegionLastRow(wb, sheetname, "title"), 3)

  ### Source
  openxlsx::writeData(wb, sheetname, inputHelperSource(source), 3,
                      namedRegionLastRow(wb, sheetname, "title") + 1,
                      name = paste(sheetname,"source", sep = "_"))

  ### Table of content caption
  openxlsx::writeData(wb, sheetname, toc_caption, 3,
                      namedRegionLastRow(wb, sheetname, "source"),
                      name = paste(sheetname,"toc", sep = "_"))
  openxlsx::addStyle(wb, sheetname, subtitleStyle(),
                     namedRegionLastRow(wb, sheetname, "toc"), 3)


  # Format ---------
  ### Set column width of first column to 1
  openxlsx::setColWidths(wb, sheetname, 1, 1)

  ### Hide gridlines
  openxlsx::showGridLines(wb, sheetname, FALSE)
}
