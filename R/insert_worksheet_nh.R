#' insert_worksheet_nh
#'
#' @description Function to add formatted worksheets without headers to an
#'  existing workbook
#' @note The function does not write the result into a .xlsx file. A separate
#'  call to `openxlsx::saveWorkbook()` is required.
#' @param data data to be included in the XLSX-table.
#' @param wb workbook object to write new worksheet in.
#' @param title title of the table and the sheet
#' @param sheetname name of the sheet-tab.
#' @param source source of the data. Defaults to "statzh".
#' @param metadata metadata-information to be included. Defaults to NA.
#' @param grouplines defaults to NA. Can be used to separate grouped variables visually.
#' @param group_names Name(s) of the second header(s).
#'  Format: List e.g list(c("title 1", "title 2", "title 3"))
#' @keywords insert_worksheet, openxlsx
#' @importFrom dplyr "%>%"
#' @importFrom purrr walk
#' @export
#' @examples
#' ## Create workbook
#' export <- openxlsx::createWorkbook()
#'
#' ## insert a new worksheet
#' insert_worksheet_nh(head(mtcars), export, "data1",
#'   title = "Title", source = "statzh", metadata = "Note: ...")
#'
#' ## insert a further worksheet
#' insert_worksheet_nh(tail(mtcars), export, "data2",
#'   title = "Title", source = "statzh", metadata = "Note: ...")
#'
#' ## insert a worksheet with group lines and second header
#' insert_worksheet_nh(data = head(mtcars), wb = export,
#'   title = "grouplines", source = "statzh", metadata = "Note: ...",
#'   grouplines = c(1,5,6), group_names = "carb")
#'
#' \dontrun{
#' ## save workbook
#'  openxlsx::saveWorkbook(export,"example.xlsx")
#' }
insert_worksheet_nh <- function(data, wb, sheetname = "Daten", title = "Title",
                                source = "statzh", metadata = NA, grouplines = NA,
                                group_names = NA){

  # Process input (substitute default values) -----
  source <- inputHelperSource(source)
  metadata <- inputHelperMetadata(metadata)


  # Determine start/end rows, row and column extents of content blocks -----
  ### Descriptives
  title_start_row <- 1
  source_start_row <- title_start_row + 1
  metadata_start_row <- source_start_row + length(source)
  metadata_end_row <- metadata_start_row + length(metadata)
  # descriptives_row_extent <- title_start_row:metadata_end_row

  ### Data
  data_start_row <- metadata_end_row + 3
  data_end_row <- data_start_row + nrow(data)
  # data_row_extent <- data_start_row:data_end_row


  # Initialize new worksheet ------
  sheetname <- verifyInputSheetname(sheetname)
  openxlsx::addWorksheet(wb, sheetname)


  # Insert title, metadata, and sources into worksheet --------
  ### Title
  openxlsx::writeData(wb, sheetname, title)
  openxlsx::addStyle(wb, sheetname, style_title(), 1, 1)

  ### Source
  openxlsx::writeData(wb, sheetname, source, startRow = source_start_row)

  ### Metadata
  openxlsx::writeData(wb, sheetname, metadata, startRow = metadata_start_row)

  ### Merge cells with title, metadata, and sources to ensure that they're displayed properly
  purrr::walk(1:metadata_end_row, ~openxlsx::mergeCells(wb, sheetname, 1:26, rows = .))


  # Insert data --------
  ### Insert second header
  if (!any(is.na(group_names))){
    insert_second_header(wb, sheetname, data_start_row, group_names, grouplines, data)
    data_start_row <- data_start_row + 1
  }

  ### Pad colnames using whitespaces for better auto-fitting of column width
  colnames(data) <- paste0(colnames(data), "  ", sep = "")

  ### Write data after checking for leftover grouping
  openxlsx::writeData(wb, sheetname, verifyDataUngrouped(data), startRow = data_start_row, rowNames = FALSE, withFilter = FALSE)
  openxlsx::addStyle(wb, sheetname, style_header(), data_start_row, 1:ncol(data), gridExpand = TRUE, stack = TRUE)


  # Grouplines ---------
  if (any(!is.na(grouplines))){
    if (!any(is.na(group_names))){
      data_start_row <- data_start_row - 1
      data_end_row <- nrow(data) + data_start_row + 1

    } else {
      data_end_row <- nrow(data) + data_start_row
    }

    if (is.numeric(grouplines)){
      groupline_numbers <- grouplines

    } else if (is.character(grouplines)){
      groupline_numbers <- get_groupline_index_by_pattern(grouplines, data)
    }

    openxlsx::addStyle(wb, sheetname, style_leftline(), data_start_row:data_end_row, groupline_numbers, gridExpand = TRUE, stack = TRUE)
  }

  # Format --------

  ### Define minimum column width
  options("openxlsx.minWidth" = 5)

  ### Use automatic column width for columns with data
  openxlsx::setColWidths(wb, sheetname, 1:ncol(data), "auto", ignoreMergedCells = TRUE)
}
