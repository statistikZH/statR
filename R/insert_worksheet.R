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
  contact_start_col <- max(ncol(data) - 2, 4)
  contact_end_col <- 26

  openxlsx::writeData(wb, sheetname, inputHelperContactInfo(contactdetails),
                      contact_start_col, 2,
                      name = paste(sheetname,"contact", sep = "_"))
  openxlsx::writeData(wb, sheetname, inputHelperHomepage(homepage),
                      contact_start_col,
                      startRow = getNamedRegionLastRow(wb, sheetname, "contact") + 1,
                      name = paste(sheetname,"homepage", sep = "_"))

  ### Creation date
  infostring <- paste(inputHelperDateCreated(prefix = date_prefix),
                      inputHelperAuthorName(author, prefix = author_prefix))
  openxlsx::writeData(wb, sheetname, infostring, contact_start_col,
                      startRow = getNamedRegionLastRow(wb, sheetname, "homepage") + 1,
                      name = paste(sheetname, "info", sep = "_"))


  ### Horizontally merge cells to ensure that contact entries are displayed properly
  info_extent <- getNamedRegionExtent(wb, sheetname, c("contact", "homepage", "info"))
  purrr::walk(info_extent$rows,
              ~openxlsx::mergeCells(wb, sheetname, min(info_extent$col):contact_end_col, rows = .))


  ### Insert headerline after contacts ------
  openxlsx::addStyle(wb, sheetname, style_headerline(),
                     rows = getNamedRegionLastRow(wb, sheetname, "info"),
                     cols = 1:ncol(data), gridExpand = TRUE, stack = TRUE)


  # Insert descriptives into worksheet --------
  ### Title
  openxlsx::writeData(wb, sheetname, title,
                      startRow = getNamedRegionLastRow(wb, sheetname, "info") + 3,
                      name = paste(sheetname, "title", sep = "_"))
  openxlsx::addStyle(wb, sheetname, style_title(),
                     getNamedRegionLastRow(wb, sheetname, "title"), 1,
                     gridExpand = TRUE)

  ### Source
  openxlsx::writeData(wb, sheetname, inputHelperSource(source),
                      startRow = getNamedRegionLastRow(wb, sheetname, "title") + 1,
                      name = paste(sheetname,"source", sep = "_"))

  ### Metadata
  openxlsx::writeData(wb, sheetname, inputHelperMetadata(metadata),
                      startRow = getNamedRegionLastRow(wb, sheetname, "source") + 1,
                      name = paste(sheetname,"metadata", sep = "_"))

  ### Horizontally merge cells containing descriptive info to ensure that they're displayed properly
  descriptives_extent <- getNamedRegionExtent(wb, sheetname, c("title","source","metadata"))
  purrr::walk(descriptives_extent$row,
              ~openxlsx::mergeCells(wb, sheetname, 1:contact_end_col, rows = .))


  # Insert data --------
  ### Pad colnames using whitespaces for better auto-fitting of column width
  colnames(data) <- paste0(colnames(data), "  ", sep = "")

  ### Write data
  openxlsx::writeData(wb, sheetname, verifyDataUngrouped(data),
                      startRow = getNamedRegionLastRow(wb, sheetname, "metadata") + 3,
                      rowNames = FALSE, withFilter = FALSE,
                      name = paste(sheetname,"data", sep = "_"))

  openxlsx::addStyle(wb, sheetname, style_header(),
                     rows = getNamedRegionFirstRow(wb, sheetname, "data"),
                     cols = 1:ncol(data),
                     gridExpand = TRUE, stack = TRUE)

  ### Grouplines
  if (!is.null(grouplines)){
    openxlsx::addStyle(wb, sheetname, style_leftline(),
                       getNamedRegionExtent(wb, sheetname, "data")$row,
                       grouplines, gridExpand = TRUE, stack = TRUE)
  }

  ### Define minimum column width
  options("openxlsx.minWidth" = 5)

  ### Use automatic columnwidth for columns with data
  openxlsx::setColWidths(wb, sheetname, 1:ncol(data), "auto", ignoreMergedCells = TRUE)
}
