#' splitXLSX()
#'
#' @description Function to export data from R as formatted .xlsx-file and spread
#'  them over several worksheets based on a grouping variable (e.g., year).
#' @note User should make sure that the grouping variable is of binary,
#'   categorical or other types with a limited number of levels.
#' @inheritParams insert_worksheet
#' @param file file name of the output .xlsx-file. The extension is added
#'  automatically.
#' @param sheetvar name of the variable used to split the data and spread them
#'  over several sheets.
#' @examples
#'splitXLSX(data = mtcars,
#'          title = "Motor trend car road tests",
#'          file = tempfile(fileext = ".xlsx"),
#'          sheetvar = cyl,
#'          source = paste("Source: Henderson and Velleman (1981),",
#'                         "Building multiple regression models interactively.",
#'                         "Biometrics, 37, 391–411."),
#'          metadata = paste("The data was extracted from the 1974 Motor Trend",
#'                           "US magazine and comprises fuel consumption and",
#'                           "10 aspects of automobile design and performance",
#'                           "for 32 automobiles (1973–74 models)."))
#' @keywords splitXLSX
#' @export
splitXLSX <- function(data,
                      file,
                      sheetvar,
                      title = "Titel",
                      source = getOption("statR_source"),
                      metadata = NA,
                      logo = getOption("statR_logo"),
                      contactdetails = inputHelperContactInfo(compact = TRUE),
                      homepage = getOption("statR_homepage"),
                      author = "user",
                      grouplines = NA,
                      group_names = NA) {

  # create workbook ------
  wb <- openxlsx::createWorkbook()

  # get values of the variable that is used to split the data -------
  data <- as.data.frame(data)
  sheetvalues <- unique(data[, c(deparse(substitute(sheetvar)))])
  col_name <- rlang::enquo(sheetvar)

  # loop to split values of the variable used to split the data -----
  for (sheetvalue in sheetvalues){
    sheettitle <- paste0(title, " (", deparse(substitute(sheetvar)), ": ", sheetvalue, ")")
    data_subset <- verifyDataUngrouped(data) %>%
      dplyr::filter((!!col_name) == sheetvalue)

    sheetname <- paste(deparse(substitute(sheetvar)), sheetvalue, sep = "_")
    insert_worksheet(wb, sheetname = sheetname, data = data_subset,
                     title = sheettitle, source = source, metadata = metadata,
                     logo = logo, contactdetails = contactdetails,
                     homepage = homepage, author = author,
                     grouplines = grouplines, group_names = group_names)
  }

  # Reverse order
  openxlsx::worksheetOrder(wb) <- rev(openxlsx::worksheetOrder(wb))

  # Save xlsx ------
  openxlsx::saveWorkbook(wb, verifyInputFilename(file), overwrite = TRUE)
}
