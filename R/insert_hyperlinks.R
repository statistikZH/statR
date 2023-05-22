#' insert_hyperlink()
#'
#' @description Function for inserting hyperlinks within an openxlsx Workbook
#' @param wb Workbook object
#' @param sheetname Name of sheet
#' @param title Title
#' @param sheet_row Sheet row
#' @param index_sheet_name Name of sheet where hyperlink should be created
#' @keywords internal
insert_hyperlink <- function(wb, sheetname, title, sheet_row, index_sheet_name){
  openxlsx::writeData(wb, index_sheet_name, title, xy = c("C", sheet_row))
  openxlsx::addStyle(wb, index_sheet_name, hyperlinkStyle(), sheet_row, cols = 3)
  openxlsx::mergeCells(wb, index_sheet_name, rows = sheet_row, cols = 3:8)

  # Set up hyperlink -------
  worksheet <- wb$sheetOrder[1]

  field_t <- wb$worksheets[[worksheet]]$sheet_data$t
  field_t[length(field_t)] <- 3

  field_v <- wb$worksheets[[worksheet]]$sheet_data$v
  field_v[length(field_v)] <- NA

  field_f <- wb$worksheets[[worksheet]]$sheet_data$f
  field_f[length(field_f)] <- paste0("<f>=HYPERLINK(&quot;#&apos;", sheetname,
                                     "&apos;!A1&quot;, &quot;", title, "&quot;)</f>")

  wb$worksheets[[worksheet]]$sheet_data$t <- as.integer(field_t)
  wb$worksheets[[worksheet]]$sheet_data$v <- field_v
  wb$worksheets[[worksheet]]$sheet_data$f <- field_f
}

#' insert_hyperlinks()
#'
#' @description Function for inserting hyperlinks within an openxlsx Workbook
#' @param wb Workbook object
#' @param sheetnames Names of sheets to create hyperlinks to
#' @param titles Titles of Hyperlinks
#' @param index_sheet_name Name of sheet where hyperlinks should be created
#' @param sheet_start_row Sheet row
#' @export
#' @keywords insert_hyperlinks
insert_hyperlinks <- function(wb, sheetnames, titles, index_sheet_name = "Index",
                              sheet_start_row = 15){

  sheet_rows <- sheet_start_row + seq(0, length(sheetnames) - 1)

  data.frame(sheetnames = sheetnames,
             titles = titles,
             sheet_row = sheet_rows) %>%
    purrr::pwalk(~insert_hyperlink(wb,
                                   sheetname = ..1,
                                   title = ..2,
                                   sheet_row = ..3,
                                   index_sheet_name = index_sheet_name))
}

