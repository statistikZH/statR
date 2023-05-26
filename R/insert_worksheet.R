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
  logo <- inputHelperLogoPath(logo)
  if (!is.null(logo)){
    openxlsx::insertImage(wb, sheetname, logo, 2.145, 0.7865, units = "in")
  }


  # Insert contact info, date created, and author -----
  ### Contact info
  openxlsx::writeData(wb, sheetname, inputHelperContactInfo(contactdetails),
                      contact_start_col, contact_start_row,
                      name = "contact")
  openxlsx::writeData(wb, sheetname, inputHelperHomepage(homepage),
                      contact_start_col,
                      startRow = getNamedRegionLastRow(wb, sheetname, "contact") + 1,
                      name = "homepage")

  ### Creation date
  infostring <- paste(inputHelperDateCreated(prefix = date_prefix),
                      inputHelperAuthorName(author, prefix = author_prefix))
  openxlsx::writeData(wb, sheetname, infostring, contact_start_col,
                      startRow = getNamedRegionLastRow(wb, sheetname, "homepage") + 1,
                      name = "info")


  ### Horizontally merge cells to ensure that contact entries are displayed properly
  info_extent <- getNamedRegionExtent(wb, sheetname, c("contact", "homepage", "info"))
  purrr::walk(info_extent$rows,
              ~openxlsx::mergeCells(wb, sheetname, min(info_extent$col):contact_end_col, rows = .))


  ### Insert headerline after contacts ------
  openxlsx::addStyle(wb, sheetname, style_headerline(),
                     rows = getNamedRegionLastRow(wb, sheetname, "homepage") + 1,
                     cols = 1:ncol(data), gridExpand = TRUE, stack = TRUE)


  # Insert descriptives into worksheet --------
  ### Title
  openxlsx::writeData(wb, sheetname, title,
                      startRow = getNamedRegionLastRow(wb, sheetname) + 3,
                      name = "title")
  openxlsx::addStyle(wb, sheetname, style_title(),
                     getNamedRegionLastRow(wb, sheetname, "title"), 1,
                     gridExpand = TRUE)

  ### Source
  openxlsx::writeData(wb, sheetname, source,
                      startRow = getNamedRegionLastRow(wb, sheetname, "title") + 1,
                      name = "source")

  ### Metadata
  openxlsx::writeData(wb, sheetname, metadata,
                      startRow = getNamedRegionLastRow(wb, sheetname, "source") + 1,
                      name = "metadata")

  ### Horizontally merge cells containing descriptive info to ensure that they're displayed properly
  descriptives_extent <- getNamedRegionExtent(wb, sheetname, c("title","source","metadata"))
  purrr::walk(descriptives_extent$row,
              ~openxlsx::mergeCells(wb, sheetname, 1:contact_end_col, rows = .))


  # Insert data --------
  ### Pad colnames using whitespaces for better auto-fitting of column width
  colnames(data) <- paste0(colnames(data), "  ", sep = "")

  ### Write data
  openxlsx::writeData(wb, sheetname, verifyDataUngrouped(data),
                      startRow = getNamedRegionLastRow(wb, sheetname) + 3,
                      rowNames = FALSE, withFilter = FALSE,
                      name = "data_region")

  openxlsx::addStyle(wb, sheetname, style_header(),
                     rows = getNamedRegionFirstRow(wb, sheetname, "data_region"),
                     cols = 1:ncol(data),
                     gridExpand = TRUE, stack = TRUE)

  ### Grouplines
  if (!is.null(grouplines)){
    openxlsx::addStyle(wb, sheetname, style_leftline(), data_row_extent, grouplines, gridExpand = TRUE, stack = TRUE)
  }

  ### Define minimum column width
  options("openxlsx.minWidth" = 5)

  ### Use automatic columnwidth for columns with data
  openxlsx::setColWidths(wb, sheetname, 1:ncol(data), "auto", ignoreMergedCells = TRUE)
}
