#' quickXLSX()
#'
#' @description Function to export data from R to a formatted .xlsx-spreadsheet.
#' @param data data to be exported.
#' @param title title to be put above the data in the worksheet.
#' @param file file name of the xlsx-file. The extension ".xlsx" is added
#'  automatically.
#' @template shared_parameters
#' @keywords quickXLSX
#' @export
#' @examples
#'quickXLSX(data = mtcars,
#'          title = "Motor trend car road tests",
#'          file = "motor_trend_car_road_tests",
#'          source = "Source: Henderson and Velleman (1981). Building multiple
#'           regression models interactively. Biometrics, 37, 391–411.",
#'          metadata = c("The data was extracted from the 1974 Motor Trend US
#'           magazine and comprises fuel consumption and 10 aspects of automobile
#'           design and performance for 32 automobiles (1973–74 models)."),
#'          contactdetails = "statzh",
#'          grouplines = FALSE,
#'          logo = "statzh",
#'          author = "user")
quickXLSX <- function(data = NA, file, title = "Title", source = "statzh",
                      metadata = NA, logo = "statzh", grouplines = FALSE,
                      contactdetails = "statzh", author = "user"){

  # Create workbook --------
  wb <- openxlsx::createWorkbook()


  # Insert data --------
  insert_worksheet(data, wb, "Inhalt", title = title, source = source,
                   metadata = metadata, logo = logo, grouplines = grouplines,
                   contactdetails = contactdetails, author = author)


  # Save workbook---------
  openxlsx::saveWorkbook(wb, verifyInputFilename(file), overwrite = TRUE)
}

