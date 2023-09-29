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
  # wb <- openxlsx::createWorkbook()

  if (length(unique(data[,sheetvar])) > 10) {
    stop("Temporary block: too many distinct values")
  }

  datasets <- split.data.frame(data, data[,sheetvar])
  sheetnames <- paste0(sheetvar, "_", names(datasets))
  titles <- paste0(title, " (", sheetvar, ": ", names(datasets), ")")
  sources <- list(source)[rep(1, length(datasets))]
  metadata <- list(metadata)[rep(1, length(datasets))]
  grouplines <- list(grouplines)[rep(1, length(datasets))]
  group_names <- list(group_names)[rep(1, length(datasets))]


  datasetsXLSX(
    file = file, datasets = datasets, sheetnames = sheetnames,
    titles = titles, sources = sources, metadata = metadata,
    grouplines = grouplines, group_names = group_names, overwrite = TRUE)
}


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
aXLSX <- function(
    data, file, title = "Title", source = getOption("statR_source"),
    metadata = NA, logo = getOption("statR_logo"), contactdetails = inputHelperContactInfo(),
    author = "user", grouplines = NA, group_names = NA) {

  # Initialize Workbook object -------
  wb <- openxlsx::createWorkbook()

  # Insert data -----
  insert_worksheet_nh(
    wb, data = data, title = title, source = source, metadata = NA,
    grouplines = grouplines, group_names = group_names)

  # Insert metadata -------
  insert_metadata_sheet(
    wb, title = title, source = source, metadata = metadata, logo = logo,
    contactdetails = contactdetails, author = author)

  # Write workbook to disk --------
  openxlsx::saveWorkbook(wb, verifyInputFilename(file), overwrite = TRUE)
}


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
quickXLSX <- function(
    data = NA, file, title = "Title", source = getOption("statR_source"),
    metadata = NA, logo = getOption("statR_logo"),
    contactdetails = inputHelperContactInfo(compact = TRUE),
    author = "user", grouplines = NA, group_names = NA) {


  # Create workbook --------
  wb <- openxlsx::createWorkbook()

  # Insert data --------
  insert_worksheet(
    wb, sheetname = "Inhalt", data = data, title = title, source = source,
    metadata = metadata, logo = logo, contactdetails = contactdetails, author = author,
    grouplines = grouplines, group_names = group_names)


  # Save workbook---------
  openxlsx::saveWorkbook(wb, verifyInputFilename(file), overwrite = TRUE)
}
