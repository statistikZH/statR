# quickXLSX: Function to create formatted spreadsheet with a single worksheet automatically

#' quickXLSX
#'
#' Function to create a formated single-worksheet XLSX automatically
#' @param data data to be included in the XLSX-table.
#' @param name title of the table.
#' @param filename filename of the xlsx-file. Defaults to "file".
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
#' quickXLSX(head(mtcars), "mtcars22", filename="testfile22", source="alternative source",contactdetails="blavla", grouplines = c(1,2,3), metadata = "remarks:",logo="L:/STAT/08_DS/06_Diffusion/Logos_Bilder/LOGOS/STAT_LOGOS/nacht_map.png")


quickXLSX <-function (data, name, filename="file", metadata = NA, logo = "statzh",
                      contactdetails="statzh",source="statzh",grouplines=FALSE) {

  #create workbook - eventually cut string after certain position with substr
  wb <- openxlsx::createWorkbook(paste(name))


insert_worksheet(data, wb, sheetname = name, source, contactdetails, metadata, logo,  grouplines=grouplines)


  #save workbook
  openxlsx::saveWorkbook(wb, paste(substr(filename,0,20), ".xlsx", sep = ""),
                         overwrite = TRUE)
}

