#' insert_second_header()
#'
#' Function to export data from R to a formatted .xlsx-file.
#'
#' The data is exported
#' to the first sheet. Metadata information is exported to the second sheet.
#'
#' @param data data to be exported.
#' @param file file name of the xlsx-file. The extension ".xlsx" is added automatically.
#' @param title title to be put above the data in the worksheet.
#' @template shared_parameters
#' @keywords aXLSX
#' @examples
#' \donttest{
#' \dontrun{
#' # Beispiel anhand des Datensatzes 'mtcars'
#'dat <- mtcars
#'
#'insert_second_header(wb,
#' sheetname,
#' data_start_row,
#' group_names,
#' grouplines,
#' data)
#' }
#' }
insert_second_header <- function(wb, sheetname, data_start_row, group_names,
                                 grouplines, data){

  if (is.character(grouplines)){
    groupline_numbers <- get_groupline_index_by_pattern(grouplines, data)

  } else if (is.numeric(grouplines)){
    groupline_numbers <- grouplines
  }

  openxlsx::addStyle(wb,
                     sheet = sheetname,
                     style = style_header(),
                     rows = data_start_row,
                     cols = 1:ncol(data)
  )

  purrr::walk2(groupline_numbers,
               group_names,
               ~openxlsx::writeData(wb,
                                    sheet = sheetname,
                                    x = .y,
                                    startCol = .x,
                                    colNames = F,
                                    startRow = data_start_row)
  )
}