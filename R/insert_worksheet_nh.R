#' insert_worksheet_nh
#'
#' @description Function to add formatted worksheets to an existing Workbook.
#'   The worksheets do not include contact information or logos.
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
insert_worksheet_nh <- function(data,
                                wb,
                                sheetname = "Daten",
                                title = "Title",
                                source = "statzh",
                                metadata = NA,
                                grouplines = NA,
                                group_names = NA){
  # Initialize new worksheet ------
  sheetname <- verifyInputSheetname(sheetname)

  if (!(sheetname %in% names(wb))){
    openxlsx::addWorksheet(wb, sheetname)
    start_row <- 1

  } else {
    start_row <- getNamedRegionLastRow(wb, sheetname) + 3
  }


  # Insert title, metadata, and sources into worksheet --------
  ### Title
  openxlsx::writeData(wb, sheetname, title, startCol = 1, startRow = start_row,
                      name = paste(sheetname, "title", sep = "_"))
  openxlsx::addStyle(wb, sheetname, style_title(), start_row, 1)

  ### Source
  openxlsx::writeData(wb, sheetname, inputHelperSource(source),
                      startRow = getNamedRegionLastRow(wb, sheetname, "title") + 1,
                      name = paste(sheetname,"source", sep = "_"))
  openxlsx::addStyle(wb, sheetname, style_subtitle(),
                     getNamedRegionLastRow(wb, sheetname, "source"), 1,
                     stack = TRUE, gridExpand = TRUE)

  ### Metadata
  openxlsx::writeData(wb, sheetname, inputHelperMetadata(metadata),
                      startRow = getNamedRegionLastRow(wb, sheetname, "source") + 1,
                      name = paste(sheetname,"metadata", sep = "_"))
  openxlsx::addStyle(wb, sheetname, style_subtitle(),
                     getNamedRegionLastRow(wb, sheetname, "metadata"), 1,
                     stack = TRUE, gridExpand = TRUE)

  ### Merge cells with title, metadata, and sources to ensure that they're displayed properly
  descr_extent <- getNamedRegionExtent(wb, sheetname, c("title", "source", "metadata"))
  purrr::walk(descr_extent$row, ~openxlsx::mergeCells(wb, sheetname, 1:26, rows = .))

  ### Add Line wrapping
  openxlsx::addStyle(wb, sheetname, style_wrap(), descr_extent$row, 1,
                     stack = TRUE, gridExpand = TRUE)


  # Insert data --------
  data_start_row <- getNamedRegionLastRow(wb, sheetname, "metadata") + 3


  # Grouplines ---------
  if (!any(is.null(grouplines)) & !any(is.na(grouplines))){
    if (is.numeric(grouplines)){
      groupline_numbers <- grouplines

    } else if (is.character(grouplines)){
      groupline_numbers <- get_groupline_index_by_pattern(grouplines, data)
    }

    ### Insert second header
    if (!any(is.null(group_names)) & !any(is.na(group_names))){
      insert_second_header(wb, sheetname, data_start_row, group_names, grouplines, data)
      data_start_row <- data_start_row + 1
    }

    data_row_extent <- data_start_row + 1:nrow(data) - 1
    openxlsx::addStyle(wb, sheetname, style_leftline(),
                       data_row_extent, groupline_numbers,
                       gridExpand = TRUE, stack = TRUE)
  }

  ### Pad colnames using whitespaces for better auto-fitting of column width
  colnames(data) <- paste0(colnames(data), "  ", sep = "")

  ### Write data after checking for leftover grouping
  openxlsx::writeData(wb, sheetname, verifyDataUngrouped(data),
                      startRow = data_start_row, rowNames = FALSE,
                      withFilter = FALSE, name = paste(sheetname,"data", sep = "_"))
  openxlsx::addStyle(wb, sheetname, style_header(),
                     data_start_row, 1:ncol(data),
                     gridExpand = TRUE, stack = TRUE)

  # Format --------
  ### Define minimum column width
  options("openxlsx.minWidth" = 5)

  ### Use automatic column width for columns with data
  openxlsx::setColWidths(wb, sheetname, 1:ncol(data), "auto", ignoreMergedCells = TRUE)
}
