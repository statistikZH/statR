#' insert_second_header()
#'
#' @description Function to add a second header row
#' @inheritParams insert_worksheet
#' @param data_start_row Row index for first row with data
#' @keywords internal
insert_second_header <- function(wb, sheetname, data_start_row, group_names,
                                 grouplines, data) {

  if (is.character(grouplines)) {
    groupline_numbers <- match(grouplines, colnames(mtcars))

  } else if (is.numeric(grouplines)) {
    groupline_numbers <- grouplines
  }

  # Write data -------
  purrr::walk2(groupline_numbers, group_names,
               ~openxlsx::writeData(wb, sheet = sheetname, x = .y,
                                    startCol = .x, startRow = data_start_row))

  # # Apply style ---------
  openxlsx::addStyle(
    wb, sheet = sheetname, style = style_header(), rows = data_start_row,
    cols = 1:ncol(data))

  # Merge cells for second header
  groupline_numbers <- unique(c(sort(groupline_numbers), ncol(data) + 1))
  purrr::walk2(
    head(groupline_numbers, -1), tail(groupline_numbers, -1) - 1,
    ~openxlsx::mergeCells(wb, sheetname, .x:.y, data_start_row))

  purrr::walk(head(groupline_numbers, - 1),
              ~openxlsx::addStyle(wb, sheet = sheetname, style = style_leftline(),
                                  rows = data_start_row, cols = ., stack = TRUE))
  # Center second header
  purrr::walk2(
    head(groupline_numbers, -1), tail(groupline_numbers, -1) - 1,
    ~openxlsx::addStyle(
      wb, sheet = sheetname, style = openxlsx::createStyle(halign = "center"),
      rows = data_start_row, cols = .x:.y, stack = TRUE))

}

