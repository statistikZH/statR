# statXLSX: Function to create formatted with multiple worksheets spreadsheets automatically

#' statXLSX
#'
#' Function to create formatted spreadsheets automatically
#' @param data data to be included in the XLSX-table.
#' @param name title of the table / file.
#' @param filename filename of the xlsx-file. Defaults to "file".
#' @param sheetvar variable which contains the variable to be used to split the data across several sheets.
#' @param source source of the data. Defaults to "statzh".
#' @param metadata metadata-information to be included. Defaults to NA.
#' @param logo path of the file to be included as logo (png / jpeg / svg). Defaults to "statzh"
#' @param contactdetails contactdetails of the data publisher. Defaults to "statzh".
#' @param grouplines columns to be separated visually by vertical lines.
#' @keywords statXLSX
#' @export
#' @examples
#' # Generation of a spreadsheet with four worksheets (one per 'carb'-category).
#' # Can be used to generate worksheets for multiple years.
#'
#' statXLSX(mtcars[c(1:10),],"mtcars",carb,grouplines=c(1,5,6))
#'
#' statXLSX(head(mtcars),"mpgcars", carb, filename="testfile33",grouplines = c(1,2,3), metadata = "remarks: the source of the data is ...",source="canton of zurich",logo="L:/STAT/08_DS/06_Diffusion/Logos_Bilder/LOGOS/STAT_LOGOS/nacht_map.png")

# Function

statXLSX <- function(data, name, filename="file", sheetvar, grouplines = FALSE, metadata = NA, logo = "statzh",
                     contactdetails="statzh",source="statzh") {


  #extrahiere colname
  col_name <- rlang::enquo(sheetvar)

  # create workbook
  wb <- openxlsx::createWorkbook(filename)

  ### Loop for multiple worksheets ------------


  sheetvalues <- unique(data[,c(deparse(substitute(sheetvar)))])


  for (sheetvalue in sheetvalues){

    #get index

    insert_worksheet(as.data.frame(data %>% dplyr::filter((!!col_name) == sheetvalue)%>%ungroup()),
                                   wb, sheetname = paste(name, sheetvalue),
                                   metadata, logo, grouplines, contactdetails,source,grouplines=grouplines)


  }

  # --------------

  openxlsx::worksheetOrder(wb)<-rev(openxlsx::worksheetOrder(wb))


  openxlsx::saveWorkbook(wb, paste(filename,".xlsx",sep=""), overwrite = TRUE)

}



