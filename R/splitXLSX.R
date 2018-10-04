# splitXLSX: Function to create formatted with multiple worksheets spreadsheets automatically

#' splitXLSX
#'
#' Function to create formatted spreadsheets automatically
#' @param data data to be included in the XLSX-table.
#' @param file filename of the xlsx-file. No Default.
#' @param name title of the table in the worksheet, defaults to "Titel" + the value of the variable used to split the dataset across sheets.
#' @param sheetvar variable which contains the variable to be used to split the data across several sheets.
#' @param source source of the data. Defaults to "statzh".
#' @param metadata metadata-information to be included. Defaults to NA.
#' @param logo path of the file to be included as logo (png / jpeg / svg). Defaults to "statzh"
#' @param contactdetails contactdetails of the data publisher. Defaults to "statzh".
#' @param grouplines columns to be separated visually by vertical lines.
#' @keywords splitXLSX
#' @export
#' @examples
#' # Generation of a spreadsheet with four worksheets (one per 'carb'-category).
#' # Can be used to generate worksheets for multiple years.
#'
#' splitXLSX(mtcars[c(1:10),],"mtcars",carb,grouplines=c(1,5,6))
#'
#' splitXLSX(head(mtcars),carb, file="filename",grouplines = c(1,2,3), metadata = "remarks: ....",source="canton of zurich",logo="L:/STAT/08_DS/06_Diffusion/Logos_Bilder/LOGOS/STAT_LOGOS/nacht_map.png")

# Function

splitXLSX <- function(data, file, sheetvar, ...) {

  data <- as.data.frame(data)

  #extract colname
  col_name <- rlang::enquo(sheetvar)

  # create workbook
  wb <- openxlsx::createWorkbook(file)



  #get values of the variable that is used to split the data
  sheetvalues <- unique(data[,c(deparse(substitute(sheetvar)))])

  #Loop to split data across multiple worksheets -------

  for (sheetvalue in sheetvalues){

    #get data into workheets
    insert_worksheet(as.data.frame(data %>% dplyr::filter((!!col_name) == sheetvalue)%>%ungroup()),
                                   wb, sheetname = paste(deparse(substitute(sheetvar)), sheetvalue,sep=","), ...)


  }

  # --------------

  openxlsx::worksheetOrder(wb)<-rev(openxlsx::worksheetOrder(wb))

  #save xlsx
  openxlsx::saveWorkbook(wb, paste(file,".xlsx",sep=""), overwrite = TRUE)

  # rm(newworkbook,envir = .GlobalEnv)

}



