# quickXLSX: Function to create formatted spreadsheet with a single worksheet automatically

#' quickXLSX
#'
#' Function to create a formated single-worksheet XLSX automatically
#' @param data data to be included in the XLSX-table.
#' @param file filename of the xlsx-file. No Default.
#' @param title title of the table in the worksheet, defaults to "Titel".
#' @template shared_parameters
#' @keywords quickXLSX
#' @export
#' @examples
#' quickXLSX(data=head(mtcars), file="mtcars")
#'
#' #example with own logo, a custom filename, lines separating selected columns and some remarks
#'
#' quickXLSX(head(mtcars),
#' file="filename",
#' title="alternative title",
#' source="alternative source",
#' contactdetails="blavla",
#' grouplines = c(1,2,3),
#' metadata = "remarks:",
#' logo="L:/STAT/08_DS/06_Diffusion/Logos_Bilder/LOGOS/STAT_LOGOS/nacht_map.png")


quickXLSX <-function (data,
                      file,
                      title="Title",
                      source="statzh",
                      metadata = NA,
                      logo=NULL,
                      grouplines = FALSE,
                      contactdetails="statzh") {

  #create workbook
  wb <- openxlsx::createWorkbook(paste(file))

  #insert data
  insert_worksheet(data, wb, title=title, source=source, metadata = metadata, logo=logo, grouplines = grouplines, contactdetails=contactdetails)

  #save workbook
  openxlsx::saveWorkbook(wb, paste(file, ".xlsx", sep = ""),
                         overwrite = TRUE)


}

