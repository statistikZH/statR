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
  source = "statzh", metadata = NA, logo = "statzh", contactdetails = "statzh",
  author = "user"){


  # Process input (substitute default values) -----
  contactdetails <- inputHelperContactInfo(contactdetails)

  # Determine start/end rows of content blocks -----
  contact_start_row <- 2
  contact_end_row <- contact_start_row + length(c(contactdetails, homepage))
  title_start_row <- contact_end_row + 2
  source_start_row <- title_start_row + 1
  source_end_row <- source_start_row + length(source) + 1
  metadata_start_row <- source_end_row + 1


  # Add a new worksheet ------
  sheetname <- verifyInputSheetname(sheetname)
  openxlsx::addWorksheet(wb, sheetname)


  # Insert logo --------
  logo <- inputHelperLogoPath(logo)
  if(!is.null(logo)){
    openxlsx::insertImage(wb, sheetname, logo, 2.145, 0.7865, "in")
  }


  # Insert contact info, title, metadata, and sources into worksheet --------

  ### Contact info
  contact_info <- c(contactdetails, inputHelperHomepage(homepage))
  openxlsx::writeData(wb, sheetname, contactdetails, startCol = 12, startRow = contact_start_row)

  ### Request information
  request_info <- paste(inputHelperDateCreated(prefix = "Aktualisiert am: "),
                        inputHelperAuthorName(author, prefix = "durch: "))
  openxlsx::writeData(wb, sheetname, request_info, startCol = 12, startRow = contact_end_row + 1)

  ### Title
  openxlsx::writeData(wb, sheetname, title, startRow = title_start_row)
  openxlsx::addStyle(wb, sheetname, style_title(), rows = title_start_row, cols = 1)

  ### Source and metadata
  openxlsx::writeData(wb, sheetname, inputHelperSource(source, prefix = "Datenquelle:"), startRow = source_start_row)
  openxlsx::writeData(wb, sheetname, inputHelperMetadata(metadata, prefix = "Hinweise:"), startRow = metadata_start_row)
  openxlsx::addStyle(wb, sheetname, style_subtitle2(), rows = c(source_start_row, metadata_start_row), cols = 1)


  # Format ----------

  ### Headerline
  openxlsx::addStyle(wb, sheetname, style_headerline(), rows = contact_end_row + 2, cols = 1:26, gridExpand = TRUE, stack = TRUE)

  ### Hide gridlines
  openxlsx::showGridLines(wb, sheetname, FALSE)
}
