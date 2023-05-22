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

  # Process input (substitute default values) -----
  logo <- inputHelperLogoPath(logo)
  contactdetails <- inputHelperContactInfo(contactdetails)
  creationdate <- inputHelperDateCreated(prefix = date_prefix)
  authorname <- inputHelperAuthorName(author, prefix = author_prefix)
  source <- inputHelperSource(source, prefix = source_prefix)
  metadata <- inputHelperMetadata(metadata, prefix = metadata_prefix)


  # Determine start/end rows of content blocks -----
  contact_start_row <- 2
  contact_end_row <- contact_start_row + length(contactdetails)
  title_start_row <- contact_end_row + 2
  source_start_row <- title_start_row + 1
  source_end_row <- source_start_row + length(source) + 1
  metadata_start_row <- source_end_row + 1


  # Add a new worksheet ------
  sheetname <- verifyInputSheetname(sheetname)
  openxlsx::addWorksheet(wb, sheetname)


  # Insert logo --------
  if(!is.null(logo)){
    openxlsx::insertImage(wb, sheetname, logo, 2.145, 0.7865, units = "in")
  }


  # Insert contact info, title, metadata, and sources into worksheet --------
  ### Contact info
  openxlsx::writeData(wb, sheetname, contactdetails, 12, contact_start_row)

  ### Request information
  info_string <- paste(creationdate, authorname)
  openxlsx::writeData(wb, sheetname, info_string, 12, contact_end_row + 1)

  ### Title
  openxlsx::writeData(wb, sheetname, title, startRow = title_start_row)
  openxlsx::addStyle(wb, sheetname, style_title(), rows = title_start_row, cols = 1)

  ### Source and metadata
  openxlsx::writeData(wb, sheetname, source, startRow = source_start_row)
  openxlsx::writeData(wb, sheetname, metadata, startRow = metadata_start_row)
  openxlsx::addStyle(wb, sheetname, style_subtitle2(), c(source_start_row, metadata_start_row), 1)


  # Format ----------
  ### Headerline
  openxlsx::addStyle(wb, sheetname, style_headerline(), contact_end_row + 2, 1:26, gridExpand = TRUE, stack = TRUE)

  ### Hide gridlines
  openxlsx::showGridLines(wb, sheetname, FALSE)
}
