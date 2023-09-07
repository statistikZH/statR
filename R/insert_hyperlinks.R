#' #' insert_hyperlink()
#' #'
#' #' @description Function for inserting hyperlinks within an openxlsx Workbook
#' #' @inheritParams insert_worksheet
#' #' @param sheet_row Row where hyperlink should be inserted
#' #' @param index_sheet_name Name of sheet where hyperlink should be created
#' #' @keywords internal
#' insert_hyperlink <- function(wb, sheetname, title, sheet_row,
#'                              index_sheet_name) {
#'
#'   hyperlink_string <- makeHyperlinkString(sheetname, text = title)
#'   writeFormula(wb, index_sheet_name, hyperlink_string, startRow = sheet_row)
#'
#'
#'   # openxlsx::writeData(wb, index_sheet_name, title, startCol = 3,
  #                     startRow = sheet_row)
  # openxlsx::addStyle(wb, index_sheet_name, hyperlinkStyle(), sheet_row,
  #                    cols = 3)
  # openxlsx::mergeCells(wb, index_sheet_name, rows = sheet_row, cols = 3:8)
  #
  # # Set up hyperlink -------
  # worksheet <- wb$sheetOrder[1]
  #
  # field_t <- wb$worksheets[[worksheet]]$sheet_data$t
  # field_t[length(field_t)] <- 3
  #
  # field_v <- wb$worksheets[[worksheet]]$sheet_data$v
  # field_v[length(field_v)] <- NA
  #
  # field_f <- wb$worksheets[[worksheet]]$sheet_data$f
  # field_f[length(field_f)] <- paste0("<f>=HYPERLINK(&quot;#&apos;", sheetname,
  #                                    "&apos;!A1&quot;, &quot;", title,
  #                                    "&quot;)</f>")
  #
  # wb$worksheets[[worksheet]]$sheet_data$t <- as.integer(field_t)
  # wb$worksheets[[worksheet]]$sheet_data$v <- field_v
  # wb$worksheets[[worksheet]]$sheet_data$f <- field_f
# }


#' insert_hyperlinks()
#'
#' @description Function for inserting hyperlinks within an openxlsx Workbook
#' @inheritParams insert_worksheet
#' @inheritParams insert_hyperlink
#' @param sheetnames Names of sheets to create hyperlinks to
#' @param titles Titles of Hyperlinks
#' @param sheet_start_row Initial row after which hyperlinks should be created
#' @keywords insert_hyperlinks
#' @export
insert_hyperlinks <- function(wb, sheetnames, titles,
                              index_sheet_name = "Index",
                              sheet_start_row = 15) {

  hyperlink_strings <- makeHyperlinkString(sheetnames, text = titles)
  writeFormula(wb, index_sheet_name, hyperlink_strings, startCol = 3,
               startRow = sheet_start_row)
  addStyle(wb, index_sheet_name, hyperlinkStyle(),
           rows = sheet_start_row + seq_along(sheetnames) - 1, cols = 3)

}

