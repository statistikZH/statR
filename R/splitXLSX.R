#'
#' splitXLSX()
#'
#' @description Function to export data from R as formatted .xlsx-file and spread
#'  them over several worksheets based on a grouping variable (e.g., year).
#'
#' @note May be deprecated in upcoming version. User should make sure that the
#'  grouping variable is of binary, categorical or other types with a limited
#'  number of levels.
#'
#' @param data data to be exported.
#'
#' @param file file name of the output .xlsx-file. The extension is added
#'  automatically.
#'
#' @param sheetvar name of the variable used to split the data and spread them
#'  over several sheets.
#'
#' @param title title to be put above the data in the worksheet. the sheetvar
#'  subcategory is added in brackets.
#'
#' @template shared_parameters
#' @keywords splitXLSX
#' @export
#' @examples
#'splitXLSX(data = mtcars,
#'          title = "Motor trend car road tests",
#'          file = "motor_trend_car_road_tests",
#'          sheetvar = cyl,
#'          source = "Source: Henderson and Velleman (1981),
#'          Building multiple regression models interactively.
#'          Biometrics, 37, 391–411.",
#'          metadata = c("The data was extracted from the 1974
#'          Motor Trend US magazine and comprises fuel consumption and
#'          10 aspects of automobile design and performance for
#'          32 automobiles (1973–74 models)."),
#'          contactdetails = "statzh",
#'          grouplines = FALSE,
#'          logo = "statzh",
#'          author = "user")

splitXLSX <- function(data, file, sheetvar, title = "Titel", source = "statzh",
                      metadata = NA, logo = "statzh", grouplines = FALSE,
                      contactdetails = "statzh", author = "user"){
  warning("Deprecation")

  # create workbook
  wb <- openxlsx::createWorkbook()

  # get values of the variable that is used to split the data
  data <- as.data.frame(data)
  sheettitle <- paste0(title, " (", deparse(substitute(sheetvar)), ": ",
                      sheetvalue, ")")
  sheetvalues <- unique(data[, c(deparse(substitute(sheetvar)))])
  col_name <- rlang::enquo(sheetvar)

  # loop to split values of the variable used to split the data
  for (sheetvalue in sheetvalues) {
    data %>%
      dplyr::filter((!!col_name) == sheetvalue) %>%
      dplyr::ungroup() %>%
      insert_worksheet(wb, sheetname = sheetvalue, title = sheettitle,
                       source = source, metadata = metadata, logo = logo,
                       grouplines = grouplines, contactdetails = contactdetails,
                       author = author)
  }

  openxlsx::worksheetOrder(wb) <- rev(openxlsx::worksheetOrder(wb))

  #save xlsx
  openxlsx::saveWorkbook(wb, prep_filename(file), overwrite = TRUE)

}
