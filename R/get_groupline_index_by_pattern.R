#' aXLSX()
#'
#' Function to export data from R to a formatted .xlsx-file.
#'
#' The data is exported
#' to the first sheet. Metadata information is exported to the second sheet.
#'
#' @param data data to be exported.
#' @param file file name of the xlsx-file. The extension ".xlsx" is added automatically.
#' @param title title to be put above the data in the worksheet.
#' @template shared_parameters
#' @keywords aXLSX
#' @examples
#' \donttest{
#' \dontrun{
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
#'          grouplines = NA,
#'          logo = "statzh",
#'          author = "user")
#' }
#' }
get_groupline_index_by_pattern <- function(grouplines, data){

  get_lowest_col <- function(groupline, data){
    groupline_numbers_single <- which(grepl(groupline, names(data)))

    out <- min(groupline_numbers_single)

    return(out)
  }


  groupline_numbers <- unlist(lapply(grouplines, function(x) get_lowest_col(x, data)))

  return(groupline_numbers)
}
