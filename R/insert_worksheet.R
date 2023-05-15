#' insert_worksheet()
#'
#' @description Function to insert a formatted worksheet into an existing
#'  Workbook object.
#'
#' @note The function does not write the result into a .xlsx file.
#'  A separate call to openxlsx::saveWorkbook() is required.
#'
#' @param data data to be included.
#'
#' @param workbook workbook object to add new worksheet to.
#'
#' @param title title to be put above the data.
#'
#' @param sheetname name of the sheet tab.
#'
#' @param source source of the data. Defaults to "statzh".
#'
#' @param metadata metadata information to be included. Defaults to NA.
#'
#' @param logo path of the file to be included as logo (png / jpeg / svg).
#'  Defaults to "statzh"
#'
#' @param contactdetails contact details of the data publisher. Defaults to
#'  "statzh".
#'
#' @param grouplines defaults to FALSE. Can be used to separate grouped
#'  variables visually.
#'
#' @param author defaults to the last two letters (initials) or numbers of the
#'  internal user name.
#'
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
  sheetname <- check_sheetname(sheetname)
  source <- prep_source(source)
  metadata <- prep_metadata(metadata)
  logo <- prep_logo(logo)
  contactdetails <- prep_contact(contactdetails)
  creationdate <- prep_creationdate("Aktualisiert am")
  author <- prep_username(author)


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


  # Insert logo, contact info, title, metadata, and sources into worksheet --------

  ### Logo
  if (!is.null(logo) && file.exists(logo)){
    openxlsx::insertImage(workbook, sheet = "Inhalt", file = logo,
                          width = 2.145, height = 0.7865, units = "in")
  } else {
    message("no logo found.")
  }

  ### Contact information
  purrr::walk(contact_start_row:contact_end_row,
              ~openxlsx::mergeCells(wb, sheet = sheetname,
                                    cols = contact_start_col:contact_end_col,
                                    rows = .))

  openxlsx::writeData(wb, sheet = sheetname, x = contactdetails,
                      startCol = contact_start_col, startRow = contact_start_row,
                      headerStyle = style_wrap())


  ### Creation date
  openxlsx::writeData(wb, sheet = sheetname, x = paste(creationdate, author),
                      startCol = contact_start_col, startRow = contact_end_row + 1,
                      headerStyle = style_subtitle3())

  ### Title
  openxlsx::writeData(workbook, sheet = sheetname, x = title, startRow = title_start_row,
                      headerStyle = style_title())

  ### Metadata
  openxlsx::writeData(workbook, sheet = sheetname, x = metadata, startRow = metadata_start_row,
                      headerStyle = style_subtitle3())

  ### Source
  openxlsx::writeData(workbook, sheet = sheetname, x = source, startRow = source_start_row,
                      headerStyle = style_subtitle3())

  ### Merge cells with title, metadata, and sources to ensure that they're displayed properly
  purrr::walk(title_start_row:source_end_row,
              ~openxlsx::mergeCells(wb, sheet = sheetname,
                                    cols = 1:contact_end_col, rows = .))


  # Insert data --------

  ### Pad colnames using whitespaces for better auto-fitting of column width
  colnames(data) <- paste0(colnames(data), "  ", sep = "")

  ### Write data
  openxlsx::writeData(wb, sheet = sheetname,
                      x = as.data.frame(data %>% dplyr::ungroup()),
                      startRow = data_start_row,
                      rowNames = FALSE, withFilter = FALSE)


  # Apply styles -------

  ### Title
  openxlsx::addStyle(wb, sheet = sheetname, style = style_title(),
                     rows = title_start_row, cols = 1, gridExpand = TRUE)

  ### Data header
  openxlsx::addStyle(wb, sheet = sheetname, style = style_header(),
                     rows = data_start_row, cols = 1:ncol(data),
                     gridExpand = TRUE, stack = TRUE)

  ### Headerline
  openxlsx::addStyle(wb, sheet = sheetname, style = style_headerline(),
                     rows = contact_end_row, cols = 1:ncol(data), gridExpand = TRUE,
                     stack = TRUE)

  ### Grouplines
  if (!is.null(grouplines)){
    openxlsx::addStyle(wb, sheet = sheetname, style = style_leftline(),
                       rows = data_start_row:data_end_row, cols = grouplines,
                       gridExpand = TRUE, stack = TRUE)
  }

  ### Define minimum column width
  options("openxlsx.minWidth" = 5)

  ### Use automatic columnwidth for columns with data
  openxlsx::setColWidths(wb, sheet = sheetname, cols = 1:ncol(data),
                         widths = "auto", ignoreMergedCells = TRUE)
}
