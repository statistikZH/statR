#' aXLSX()
#'
#' Function to export data from R to a formatted .xlsx-spreadsheet.
#' @param data data to be exported.
#' @param file file name of the xlsx-file. The extension ".xlsx" is added automatically.
#' @param title title to be put above the data in the worksheet.
#' @template shared_parameters
#' @keywords aXLSX
#' @export
#' @examples
#' # Beispiel anhand des Datensatzes 'mtcars'
#'dat <- mtcars
#'
#'aXLSX(data = dat,
#'          title = "Motor trend car road tests",
#'          file = "motor_trend_car_road_tests", # '.xlsx' is automatically added
#'          source = "Source: Henderson and Velleman (1981). Building multiple
#'          regression models interactively.
#'          Biometrics, 37, 391–411.",
#'          metadata = c("The data was extracted from the 1974 Motor Trend US
#'          magazine and comprises fuel
#'          consumption and 10 aspects of automobile design and performance
#'          for 32 automobiles (1973–74 models)."),
#'          contactdetails = "statzh",
#'          grouplines = FALSE,
#'          logo = "statzh",
#'          author = "user")
#'

aXLSX <-function (data,
                      file,
                      title="Title",
                      source="statzh",
                      metadata = NA,
                      logo="statzh",
                      grouplines = FALSE,
                      contactdetails="statzh",
                      author = "user"
                      ) {

  #create workbook
  wb <- openxlsx::createWorkbook(paste(file))

  #insert data
  insert_worksheet_nh(data, wb, title=title, source=source, metadata = NA, grouplines = grouplines)

  # insert metadata
  insert_metadata_sheet(wb, title=title,
                        source=source, metadata = metadata, logo= logo,
                        contactdetails=contactdetails, author = author)

  #save workbook
  openxlsx::saveWorkbook(wb, paste(file, ".xlsx", sep = ""),
                         overwrite = TRUE)


}
