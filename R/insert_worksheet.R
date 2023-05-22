#' insert_worksheet()
#'
#' @description Function to insert a formatted worksheet into an existing
#'  Workbook object.
#' @note The function does not write the result into a .xlsx file.
#'  A separate call to openxlsx::saveWorkbook() is required.
#' @param data data to be included.
#' @param wb workbook object to add new worksheet to.
#' @param title title to be put above the data.
#' @param sheetname name of the sheet tab.
#' @param source source of the data. Defaults to "statzh".
#' @param metadata metadata information to be included. Defaults to NA.
#' @param logo path of the file to be included as logo (png / jpeg / svg).
#'  Defaults to "statzh"
#' @param contactdetails contact details of the data publisher. Defaults to
#'  "statzh".
#' @param homepage Homepage of data publisher. Defaults to "statzh".
#' @param grouplines defaults to FALSE. Can be used to separate grouped
#'  variables visually.
#' @param author defaults to the last two letters (initials) or numbers of the
#'  internal user name.
#' @param date_prefix Text shown before date
#' @param author_prefix Text shown before author name
#' @importFrom dplyr "%>%"
#' @keywords insert_worksheet
#' @export
#' @examples
#' # Initialize an empty Workbook
#' wb <- openxlsx::createWorkbook()
#'
#' insert_worksheet(data = head(mtcars),
#'                  wb = wb,
#'                  title = "mtcars",
#'                  sheetname = "carb")
#'
insert_worksheet <- function(data, wb, sheetname = "data", title = "Title",
                             source = "statzh", metadata = NA, logo = "statzh",
                             grouplines = FALSE, contactdetails = "statzh",
                             homepage = "statzh", author = "user",
                             date_prefix = "Aktualisiert am: ",
                             author_prefix = "durch: "){

  # Process input (substitute default values) -----
  logo <- inputHelperLogoPath(logo)
  contactdetails <- inputHelperContactInfo(contactdetails)
  homepage <- inputHelperHomepage(homepage)
  creationdate <- inputHelperDateCreated(prefix = date_prefix)
  authorname <- inputHelperAuthorName(author, prefix = author_prefix)
  source <- inputHelperSource(source)
  metadata <- inputHelperMetadata(metadata)


  # Determine start/end rows, row and column extents of content blocks -----
  ### Contact information
  contact_start_row <- 2
  contact_start_col <- max(ncol(data) - 2, 4)
  contact_end_col <- 26
  contact_end_row <- contact_start_row + length(contactdetails)
  contact_column_extent <- contact_start_col:contact_end_col
  contact_rows_extent <- contact_start_row:contact_end_row

  ### Descriptives
  title_start_row <- contact_end_row + 3
  source_start_row <- title_start_row + 1
  metadata_start_row <- source_start_row + length(source)
  metadata_end_row <- metadata_start_row + length(metadata)
  descriptives_column_extent <- 1:contact_end_col
  descriptives_row_extent <- title_start_row:metadata_end_row

  ### Data
  data_start_row <- metadata_end_row + 3
  data_end_row <- data_start_row + nrow(data)
  data_row_extent <- data_start_row:data_end_row


  # Initialize new worksheet ------
  sheetname <- verifyInputSheetname(sheetname)
  openxlsx::addWorksheet(wb, sheetname)


  # Insert logo ------
  if (!is.null(logo)){
    openxlsx::insertImage(wb, sheetname, logo, 2.145, 0.7865, units = "in")
  }


  # Insert contact info, date created, and author -----
  ### Contact info
  openxlsx::writeData(wb, sheetname, contactdetails, contact_start_col, contact_start_row)
  openxlsx::writeData(wb, sheetname, homepage, contact_start_col, contact_end_row)

  ### Creation date
  infostring <- paste(creationdate, authorname)
  openxlsx::writeData(wb, sheetname, infostring, contact_start_col, contact_end_row+1)


  # Insert descriptives into worksheet --------
  ### Title
  openxlsx::writeData(wb, sheetname, title, startRow = title_start_row)
  openxlsx::addStyle(wb, sheetname, style_title(), title_start_row, 1, gridExpand = TRUE)

  ### Source
  openxlsx::writeData(wb, sheetname, source, startRow = source_start_row)

  ### Metadata
  openxlsx::writeData(wb, sheetname, metadata, startRow = metadata_start_row)


  # Insert data --------
  ### Pad colnames using whitespaces for better auto-fitting of column width
  colnames(data) <- paste0(colnames(data), "  ", sep = "")

  ### Write data
  openxlsx::writeData(wb, sheetname, verifyDataUngrouped(data), startRow = data_start_row,
                      rowNames = FALSE, withFilter = FALSE)

  openxlsx::addStyle(wb, sheetname, style_header(), data_start_row, 1:ncol(data),
                     gridExpand = TRUE, stack = TRUE)


  # Format ---------
  ### Horizontally merge cells to ensure that contact entries are displayed properly
  purrr::walk(contact_rows_extent, ~openxlsx::mergeCells(wb, sheetname, contact_column_extent, rows = .))

  ### Insert headerline after contacts ------
  openxlsx::addStyle(wb, sheetname, style_headerline(), contact_end_row + 1, 1:ncol(data), gridExpand = TRUE, stack = TRUE)

  ### Horizontally merge cells containing descriptive info to ensure that they're displayed properly
  purrr::walk(descriptives_row_extent, ~openxlsx::mergeCells(wb, sheetname, descriptives_column_extent, rows = .))

  ### Grouplines
  if (!is.null(grouplines)){
    openxlsx::addStyle(wb, sheetname, style_leftline(), data_row_extent, grouplines, gridExpand = TRUE, stack = TRUE)
  }

  ### Define minimum column width
  options("openxlsx.minWidth" = 5)

  ### Use automatic columnwidth for columns with data
  openxlsx::setColWidths(wb, sheetname, 1:ncol(data), "auto", ignoreMergedCells = TRUE)
}
