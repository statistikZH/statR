#' insert_hyperlinks()
#'
#' Function for inserting hyperlinks within an openxlsx Workbook
#' @description
#' A short description...
#' @param wb Worksheet
#' @param sheetname Name of sheet
#' @param title Title
#' @param sheet_row Sheet row
#' @keywords internal
insert_hyperlinks <- function(wb, sheetname, title, sheet_row){
  openxlsx::writeData(wb, sheet = "Inhalt", x = title,
    xy = c("C", sheet_row))

  openxlsx::addStyle(wb, sheet = "Inhalt",
    style = hyperlinkStyle(), rows = sheet_row,
    cols = 3)

  openxlsx::mergeCells(wb, sheet = "Inhalt", cols = 3:8,
    rows = sheet_row)

  worksheet <- wb$sheetOrder[1]

  field_t <- wb$worksheets[[worksheet]]$sheet_data$t
  field_t[length(field_t)] <- 3

  field_v <- wb$worksheets[[worksheet]]$sheet_data$v
  field_v[length(field_v)] <- NA

  field_f <- wb$worksheets[[worksheet]]$sheet_data$f
  field_f[length(field_f)] <- paste0("<f>=HYPERLINK(&quot;#&apos;",
                                     sheetname,
                                     "&apos;!A1&quot;, &quot;",
                                     title,
                                     "&quot;)</f>")

  wb$worksheets[[worksheet]]$sheet_data$t <- as.integer(field_t)
  wb$worksheets[[worksheet]]$sheet_data$v <- field_v
  wb$worksheets[[worksheet]]$sheet_data$f <- field_f
}
