#' insert_worksheet()
#'
#' @description Function to insert a formatted worksheet into an existing
#'  Workbook object.
#' @note The function does not write the result into a .xlsx file.
#'  A separate call to openxlsx::saveWorkbook() is required.
#' @param data data to be included.
#' @param workbook workbook object to add new worksheet to.
#' @param title title to be put above the data.
#' @param sheetname name of the sheet tab.
#' @param source source of the data. Defaults to "statzh".
#' @param metadata metadata information to be included. Defaults to NA.
#' @param logo path of the file to be included as logo (png / jpeg / svg).
#'  Defaults to "statzh"
#' @param contactdetails contact details of the data publisher. Defaults to
#'  "statzh".
#' @param grouplines defaults to FALSE. Can be used to separate grouped
#'  variables visually.
#' @param author defaults to the last two letters (initials) or numbers of the
#'  internal user name.
#' @importFrom dplyr "%>%"
#' @keywords insert_worksheet
#' @export
#' @examples
#'
#' # Generation of a spreadsheet
#' wb <- openxlsx::createWorkbook("hello")
#'
#' insert_worksheet(data = head(mtcars), workbook = wb, title = "mtcars",
#'  sheetname = "carb")
#'
insert_worksheet <- function(data, workbook, sheetname = "data", title = "Title",
                             source = "statzh", metadata = NA, logo = "statzh",
                             grouplines = FALSE, contactdetails = "statzh",
                             author = "user"){

  # Process input (substitute default values) -----
  sheetname <- verifyInputSheetname(sheetname)
  source <- inputHelperSource(source)
  metadata <- inputHelperMetadata(metadata)
  contactdetails <- inputHelperContactInfo(contactdetails)
  creationdate <- inputHelperDateCreated(prefix = "Aktualisiert am: ")
  author <- inputHelperAuthorName(author, prefix = "durch: ")


  # Determine start/end rows of content blocks -----
  contact_start_row <- 2
  contact_end_row <- contact_start_row + length(contactdetails)
  title_start_row <- contact_end_row + 2
  metadata_start_row <- title_start_row + 1
  source_start_row <- metadata_start_row + length(metadata)
  source_end_row <- source_start_row + length(source)
  data_start_row <- source_end_row + 3
  data_end_row <- data_start_row + nrow(data)


  # Determine column boundaries -------

  contact_start_col <- max(ncol(data) - 2, 4)
  contact_end_col <- 26


  # Initialize new worksheet ------
  openxlsx::addWorksheet(workbook, sheetname)


  # Insert logo ------
  logo <- inputHelperLogoPath(logo)

  if (!is.null(logo)){
    openxlsx::insertImage(workbook, "Inhalt", logo, 2.145, 0.7865, "in")
  }


  # Insert contact info, title, metadata, and sources into worksheet --------

  ### Contact information

  openxlsx::writeData(wb, sheetname, contactdetails, contact_start_col, contact_start_row, headerStyle = style_wrap())
  purrr::walk(contact_start_row:contact_end_row,
              ~openxlsx::mergeCells(wb, sheetname, contact_start_col:contact_end_col,
                                    rows = .))

  openxlsx::addStyle(wb, sheetname, style_headerline(), contact_end_row, 1:ncol(data),
                     gridExpand = TRUE, stack = TRUE)

  ### Creation date
  openxlsx::writeData(wb, sheetname, paste(creationdate, author), contact_start_col, contact_end_row + 1,
                      headerStyle = style_subtitle3())

  ### Title
  openxlsx::writeData(workbook, sheetname, title, startRow = title_start_row)
  openxlsx::addStyle(wb, sheetname, style_title(), title_start_row, 1, gridExpand = TRUE)

  ### Metadata
  openxlsx::writeData(workbook, sheetname, metadata, startRow = metadata_start_row, headerStyle = style_subtitle3())

  ### Source
  openxlsx::writeData(workbook, sheetname, source, startRow = source_start_row, headerStyle = style_subtitle3())

  ### Merge cells with title, metadata, and sources to ensure that they're displayed properly
  purrr::walk(title_start_row:source_end_row,
              ~openxlsx::mergeCells(wb, sheetname, 1:contact_end_col, rows = .))


  # Insert data --------

  ### Pad colnames using whitespaces for better auto-fitting of column width
  colnames(data) <- paste0(colnames(data), "  ", sep = "")

  ### Write data
  openxlsx::writeData(wb, sheetname, verifyDataUngrouped(data), startRow = data_start_row,
                      rowNames = FALSE, withFilter = FALSE)

  openxlsx::addStyle(wb, sheetname, style_header(), data_start_row, 1:ncol(data),
                     gridExpand = TRUE, stack = TRUE)


  # Format ---------

  if (!is.null(grouplines)){
    openxlsx::addStyle(wb, sheetname, style_leftline(), data_start_row:data_end_row,
                       grouplines, gridExpand = TRUE, stack = TRUE)
  }

  ### Define minimum column width
  options("openxlsx.minWidth" = 5)

  ### Use automatic columnwidth for columns with data
  openxlsx::setColWidths(wb, sheetname, 1:ncol(data), "auto", ignoreMergedCells = TRUE)
}
