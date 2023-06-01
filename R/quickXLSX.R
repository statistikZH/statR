#' quickXLSX()
#'
#' @description Function to export data from R to a formatted .xlsx-spreadsheet.
#' @inheritParams insert_worksheet
#' @param data data to be exported.
#' @param file file name of the xlsx-file. The extension ".xlsx" is added
#' @keywords quickXLSX
#' @export
#' @examples
#'quickXLSX(data = mtcars,
#'          title = "Motor trend car road tests",
#'          file = tempfile(fileext = ".xlsx"),
#'          source = paste("Source: Henderson and Velleman (1981). Building",
#'                         "multiple regression models interactively.",
#'                         "Biometrics, 37, 391–411."),
#'          metadata = paste("The data was extracted from the 1974 Motor",
#'                           "Trend US magazine and comprises fuel consumption",
#'                           "and 10 aspects of automobile design and",
#'                           "performance for 32 automobiles (1973–74 models)."))
quickXLSX <- function(data = NA,
                      file,
                      title = "Title",
                      source = "statzh",
                      metadata = NA,
                      logo = "statzh",
                      contactdetails = "statzh",
                      author = "user",
                      grouplines = NA,
                      group_names = NA){

  # Create workbook --------
  wb <- openxlsx::createWorkbook()


  # Insert data --------
  insert_worksheet(wb, sheetname = "Inhalt", data = data, title = title,
                   source = source,
                   metadata = metadata, logo = logo,
                   contactdetails = contactdetails, author = author,
                   grouplines = grouplines, group_names = group_names)


  # Save workbook---------
  openxlsx::saveWorkbook(wb, verifyInputFilename(file), overwrite = TRUE)
}

