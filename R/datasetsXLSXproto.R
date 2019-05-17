# datasetsXLSX: Function to create formatted with multiple worksheets spreadsheets automatically

#' datasetsXLSX
#'
#' Function to create formatted spreadsheets automatically
#' @param datasets data to be included in the XLSX-table.
#' @param file filename of the xlsx-file. No Default.
#' @param name title of the table in the worksheet, defaults to "Titel" + the value of the variable used to split the dataset across sheets.
#' @param sheetvar variable which contains the variable to be used to split the data across several sheets.
#' @param source source of the data. Defaults to "statzh".
#' @param metadata metadata-information to be included. Defaults to NA.
#' @param logo path of the file to be included as logo (png / jpeg / svg). Defaults to "statzh"
#' @param contactdetails contactdetails of the data publisher. Defaults to "statzh".
#' @param grouplines columns to be separated visually by vertical lines.
#' @keywords datasetsXLSX
#' @export
#' @examples
#' # Generation of a spreadsheet with four worksheets (one per 'carb'-category).
#' # Can be used to generate worksheets for multiple years.
#'
#' datasetsXLSX(file="test",datasets=list(head(mtcars),head(diamonds)))
#'


# Function

datasetsXLSX <- function(file="test", datasets=list(head(mtcars),head(diamonds)),...){

  wb <- openxlsx::createWorkbook("hello")

  i<-0

  for (dataset in datasets){

    i <- i+1

    #dynamisch mit sheetvar!
    statR::insert_worksheet(data=dataset,workbook=wb,sheetname = i)


  }


  openxlsx::saveWorkbook(wb, paste(file, ".xlsx", sep = ""))

}

