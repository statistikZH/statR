#' splitXLSX
#'
#' Function to create formatted spreadsheets with multiple worksheets automatically.
#' @param data data to be included in the XLSX-table.
#' @param file filename of the xlsx-file. No Default.
#' @param sheetvar variable which contains the variable to be used to split the data across several sheets.
#' @param title title of the table in the worksheet, defaults to "Titel" + the value of the variable used to split the dataset across sheets.
#' @template shared_parameters
#' @keywords splitXLSX
#' @export
#' @examples
#' # Generation of a spreadsheet with four worksheets (one per 'carb'-category).
#' # Can be used to generate worksheets for multiple years.
#'
#' splitXLSX(head(mtcars),"mtcars",carb,grouplines=c(1,5,6))
#'
#' splitXLSX(head(mtcars),
#' sheetvar=carb,
#' title="Tabelle",
#' file="filename",
#' grouplines = c(1,2,3),
#' metadata = "remarks: ....",
#' source="canton of zurich",
#' logo="L:/STAT/08_DS/06_Diffusion/Logos_Bilder/LOGOS/STAT_LOGOS/nacht_map.png")

# Function

splitXLSX <- function (data,
                       file,
                       sheetvar,
                       title="Titel",
                       source="statzh",
                       metadata = NA,
                       logo=NULL,
                       grouplines = FALSE,
                       contactdetails="statzh")
{
  data <- as.data.frame(data)

  # extract column name
  col_name <- rlang::enquo(sheetvar)

  # create workbook
  wb <- openxlsx::createWorkbook(file)

  # get values of the variable that is used to split the data
  sheetvalues <- unique(data[, c(deparse(substitute(sheetvar)))])

  # loop to split values of the variable used to split the data
  for (sheetvalue in sheetvalues) {

    # get data into worksheets
    insert_worksheet(as.data.frame(data %>% dplyr::filter((!!col_name) ==
                                                            sheetvalue) %>% ungroup()), wb, sheetname = sheetvalue,
                     #shared params
                     title=paste(title, sheetvalue),
                     source=source,
                     metadata = metadata,
                     logo=logo,
                     grouplines = grouplines,
                     contactdetails=contactdetails)


  }

  # --------------

  openxlsx::worksheetOrder(wb)<-rev(openxlsx::worksheetOrder(wb))

  #save xlsx
  openxlsx::saveWorkbook(wb, paste(file,".xlsx",sep=""), overwrite = TRUE)

  # rm(newworkbook,envir = .GlobalEnv)

}






