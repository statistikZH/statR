#' insert_metadata_sheet()
#'
#' Function to  add formatted metadata information to an existing .xlsx-workbook.
#' @param wb workbook object to add new worksheet to.
#' @param title title to be put above the data.
#' @param sheetname name of the sheet tab.
#' @param source source of the data. Defaults to "statzh".
#' @param metadata metadata information to be included. Defaults to NA.
#' @param logo path of the file to be included as logo (png / jpeg / svg). Defaults to "statzh"
#' @param contactdetails contact details of the data publisher. Defaults to "statzh".
#' @param author defaults to the last two letters (initials) or numbers of the internal user name.
#' @param date_prefix Text shown before date
#' @param author_prefix Text shown before author name
#' @param source_prefix Text shown before sources
#' @param metadata_prefix Text shown before metadata
#' @importFrom dplyr "%>%"
#' @keywords insert_metadata_sheet, openxlsx
#' @export
#' @examples
#' # Generation of a spreadsheet
#' wb <- openxlsx::createWorkbook()
#' insert_metadata_sheet(wb, title = "Title of mtcars",
#'   metadata = c("Meta data information."))
#' \dontrun{
#' openxlsx::saveWorkbook(wb,"insert_metadata_sheet.xlsx")
#' }
insert_metadata_sheet <- function(wb, sheetname = "Metadaten", title = "Title",
                                  source = "statzh", metadata = NA,
                                  logo = "statzh", contactdetails = "statzh",
                                  author = "user", date_prefix = "Aktualisiert am: ",
                                  author_prefix = "durch: ",
                                  source_prefix = "Datenquelle:",
                                  metadata_prefix = "Hinweise:"){

  # Determine start/end rows of content blocks -----
  contact_start_row <- 2


  # Add a new worksheet ------
  sheetname <- verifyInputSheetname(sheetname)
  openxlsx::addWorksheet(wb, sheetname)


  # Insert logo --------
  logo <- inputHelperLogoPath(logo)
  if(!is.null(logo)){
    openxlsx::insertImage(wb, sheetname, logo, 2.145, 0.7865, units = "in")
  }


  # Insert contact info, title, metadata, and sources into worksheet --------
  ### Contact info
  openxlsx::writeData(wb, sheetname, inputHelperContactInfo(contactdetails),
                      12, 2, name = paste(sheetname,"contact", sep = "_"))

  ### Request information
  info_string <- paste(inputHelperDateCreated(prefix = date_prefix),
                       inputHelperAuthorName(author, prefix = author_prefix))
  openxlsx::writeData(wb, sheetname, info_string, 12,
                      startRow = getNamedRegionLastRow(wb, sheetname, "contact") + 1,
                      name = paste(sheetname,"info", sep = "_"))
  ### Headerline
  openxlsx::addStyle(wb, sheetname, style_headerline(),
                     getNamedRegionLastRow(wb, sheetname, "contact") + 1, 1:26,
                     gridExpand = TRUE, stack = TRUE)

  ### Title
  openxlsx::writeData(wb, sheetname, title,
                      startRow = getNamedRegionLastRow(wb, sheetname, "info") + 3,
                      name = paste(sheetname,"title", sep = "_"))
  openxlsx::addStyle(wb, sheetname, style_title(),
                     rows = getNamedRegionLastRow(wb, sheetname, "title"), cols = 1)

  ### Source and metadata
  openxlsx::writeData(wb, sheetname, inputHelperSource(source, prefix = source_prefix),
                      startRow = getNamedRegionLastRow(wb, sheetname, "title") + 1,
                      name = paste(sheetname,"source", sep = "_"))
  openxlsx::writeData(wb, sheetname, inputHelperMetadata(metadata, prefix = metadata_prefix),
                      startRow = getNamedRegionLastRow(wb, sheetname, "source") + 1,
                      name = paste(sheetname,"metadata", sep = "_"))
  openxlsx::addStyle(wb, sheetname, style_subtitle2(),
                     c(getNamedRegionFirstRow(wb, sheetname, "source"),
                       getNamedRegionFirstRow(wb, sheetname, "metadata")), 1)

  ### Hide gridlines
  openxlsx::showGridLines(wb, sheetname, FALSE)
}
