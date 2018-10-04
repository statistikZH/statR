# quickXLSX: Function to create formatted spreadsheet with a single worksheet automatically

#' quickXLSX
#'
#' Function to create a formated single-worksheet XLSX automatically
#' @param data data to be included in the XLSX-table.
#' @param file filename of the xlsx-file. No Default.
#' @param title title of the table in the worksheet, defaults to "Titel".
#' @param source source of the data. Defaults to "statzh".
#' @param metadata metadata-information to be included. Defaults to NA.
#' @param logo path of the file to be included as logo (png / jpeg / svg). Defaults to "statzh"
#' @param contactdetails contactdetails of the data publisher. Defaults to "statzh".
#' @param grouplines columns to be separated visually by vertical lines.
#' @keywords quickXLSX
#' @export
#' @examples
#' quickXLSX(head(mtcars), "mtcars")
#'
#' #example with own logo, a custom filename, lines separating selected columns and some remarks
#'
#' quickXLSX(head(mtcars),  file="filename", title="alternative title", source="alternative source",contactdetails="blavla", grouplines = c(1,2,3), metadata = "remarks:",logo="L:/STAT/08_DS/06_Diffusion/Logos_Bilder/LOGOS/STAT_LOGOS/nacht_map.png")


quickXLSX <-function (data, file, ...) {

  #create workbook
  wb <- openxlsx::createWorkbook(paste(file))

  #insert data
  insert_worksheet(data, wb, ...)

  #save workbook
  openxlsx::saveWorkbook(wb, paste(file, ".xlsx", sep = ""),
                         overwrite = TRUE)

  # rm(newworkbook,envir = .GlobalEnv)


}

