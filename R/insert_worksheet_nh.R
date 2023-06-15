#' insert_worksheet_nh
#'
#' @description Function to add formatted worksheets to an existing Workbook.
#'   The worksheets do not include contact information or logos.
#' @note The function does not write the result into a .xlsx file. A separate
#'  call to `openxlsx::saveWorkbook()` is required.
#' @inheritParams insert_worksheet
#' @examples
#' ## Create workbook
#' wb <- openxlsx::createWorkbook()
#'
#' ## insert a new worksheet
#' insert_worksheet_nh(wb, sheetname = "data1", data = head(mtcars),
#'   title = "Title", source = "statzh", metadata = "Note: ...")
#'
#' ## insert a further worksheet
#' insert_worksheet_nh(wb, sheetname = "data2", data = tail(mtcars),
#'   title = "Title", source = "statzh", metadata = "Note: ...")
#'
#' ## insert a worksheet with group lines and second header
#' insert_worksheet_nh(wb, sheetname = "data3", data = head(mtcars),
#'                     title = "grouplines", source = "statzh",
#'                     metadata = "Note: ...",
#'                     grouplines = c(1,5,6), group_names = "carb")
#' @keywords insert_worksheet_nh
#' @export
insert_worksheet_nh <- function(wb,
                                sheetname = "Daten",
                                data,
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
    start_row <- namedRegionLastRow(wb, sheetname) + 3
  }

  # Insert title, metadata, and sources into worksheet --------
  ### Title
  openxlsx::writeData(wb, sheetname, title, startCol = 1, startRow = start_row,
                      name = paste(sheetname, "title", sep = "_"))
  openxlsx::addStyle(wb, sheetname, style_title(), start_row, 1)

  ### Source
  openxlsx::writeData(wb, sheetname,
                      inputHelperSource(source),
                      startRow = namedRegionLastRow(wb, sheetname, "title") + 1,
                      name = paste(sheetname,"source", sep = "_"))
  openxlsx::addStyle(wb, sheetname, style_subtitle(),
                     namedRegionFirstRow(wb, sheetname, "source"), 1,
                     stack = TRUE, gridExpand = TRUE)

  ### Metadata
  openxlsx::writeData(wb, sheetname,
                      inputHelperMetadata(metadata),
                      startRow = namedRegionLastRow(wb, sheetname, "source") + 1,
                      name = paste(sheetname,"metadata", sep = "_"))
  openxlsx::addStyle(wb, sheetname, style_subtitle(),
                     namedRegionFirstRow(wb, sheetname, "metadata"), 1,
                     stack = TRUE, gridExpand = TRUE)


  ### Merge cells with title, metadata, and sources to ensure that they're displayed properly
  purrr::walk(namedRegionRowExtent(wb, sheetname, c("title", "source", "metadata")),
              ~openxlsx::mergeCells(wb, sheetname, cols = 1:26, rows = .))

  ### Add Line wrapping
  openxlsx::addStyle(wb, sheetname, style_wrap(),
                     namedRegionRowExtent(wb, sheetname, c("title", "source", "metadata")), 1,
                     stack = TRUE, gridExpand = TRUE)


  # Insert data --------
  data_start_row <- namedRegionLastRow(wb, sheetname, "metadata") + 2


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
                      withFilter = FALSE,
                      name = paste(sheetname, "data", sep = "_"))
  openxlsx::addStyle(wb, sheetname, style_header(),
                     data_start_row, 1:ncol(data),
                     gridExpand = TRUE, stack = TRUE)

  # Format --------
  ### Define minimum column width
  options("openxlsx.minWidth" = 5)

  ### Use automatic column width for columns with data
  openxlsx::setColWidths(wb, sheetname, 1:ncol(data), "auto", ignoreMergedCells = TRUE)
}
