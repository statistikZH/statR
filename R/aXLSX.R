#' aXLSX()
#'
#' @description Function to export data from R to a formatted .xlsx-file. The
#'  data is exported to the first sheet. Metadata information is exported to
#'  the second sheet.
#' @inheritParams insert_worksheet
#' @param file Path of output xlsx-file.
#' @examples
#' aXLSX(data = mtcars,
#'       title = "Motor trend car road tests",
#'       file = tempfile(fileext = ".xlsx"),
#'       source = paste("Source: Henderson and Velleman (1981). Building",
#'                      "multiple regression models interactively.",
#'                      "Biometrics, 37, 391–411."),
#'       metadata = paste("The data was extracted from the 1974",
#'                        "Motor Trend US magazine and comprises fuel",
#'                        "consumption and 10 aspects of automobile design",
#'                        "and performance for 32 automobiles",
#'                        "(1973–74 models)."))
#' @keywords aXLSX
#' @export
aXLSX <- function(data,
                  file,
                  title = "Title",
                  source = "statzh",
                  metadata = NA,
                  logo = "statzh",
                  contactdetails = "statzh",
                  author = "user",
                  grouplines = NA,
                  group_names = NA){

  # Initialize Workbook object -------
  wb <- openxlsx::createWorkbook()

  # Insert data -----
  insert_worksheet_nh(wb, data = data, title = title,
                      source = source, metadata = NA,
                      grouplines = grouplines,
                      group_names = group_names)

  # Insert metadata -------
  insert_metadata_sheet(wb, title = title, source = source,
                        metadata = metadata, logo = logo,
                        contactdetails = contactdetails,
                        author = author)

  # Write workbook to disk --------
  openxlsx::saveWorkbook(wb, verifyInputFilename(file), overwrite = TRUE)
}
