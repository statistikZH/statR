#' aXLSX()
#'
#' @description Function to export data from R to a formatted .xlsx-file. The
#'  data is exported to the first sheet. Metadata information is exported to
#'  the second sheet.
#' @param data data to be exported.
#' @param file file name of the xlsx-file. The extension ".xlsx" is added automatically.
#' @param title title to be put above the data in the worksheet.
#' @template shared_parameters
#' @keywords aXLSX
#' @export
#' @examples
#' dataset <- mtcars
#' source_string <- paste("Source: Henderson and Velleman (1981).",
#'   "Building multiple regression models interactively.",
#'   "Biometrics, 37, 391–411.")
#'
#' metadata_string <- paste("The data was extracted from the 1974",
#'   "Motor Trend US magazine and comprises fuel consumption and",
#'   "10 aspects of automobile design and performance for 32 automobiles",
#'   "(1973–74 models).")
#' \donttest{
#' \dontrun{
#' aXLSX(data = mtcars,
#'       title = "Motor trend car road tests",
#'       file = "motor_trend_car_road_tests",
#'       source = source_string,
#'       metadata = metadata_string,
#'       contactdetails = "statzh",
#'       grouplines = NA,
#'       logo = "statzh",
#'       author = "user")
#' }
#' }
aXLSX <- function(data, file, title = "Title", source = "statzh", metadata = NA,
                  logo = "statzh", grouplines = NA, contactdetails = "statzh",
                  author = "user"){

  # Initialize Workbook object -------
  wb <- openxlsx::createWorkbook()

  # Insert data -----
  insert_worksheet_nh(data, wb, title = title, source = source, metadata = NA,
                      grouplines = grouplines)

  # Insert metadata -------
  insert_metadata_sheet(wb, title = title, source = source, metadata = metadata,
                        logo = logo, contactdetails = contactdetails,
                        author = author)

  # Write workbook to disk --------
  openxlsx::saveWorkbook(wb, verifyInputFilename(file), overwrite = TRUE)
}
