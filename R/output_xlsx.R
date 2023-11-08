#'datasetsXLSX()
#'
#'@description Function to export several datasets and/or figures from R to an
#'  .xlsx-file. The function creates an overview sheet and separate sheets for
#'  each dataset/figure.
#'@details When including figures, the heights and widths need to be specified
#'  as a vector. For example, say you have one dataset and two figures that you
#'  would like to export. widths = c(5,6) then suggests that the first figure
#'  will be 5 inches wide, the second 6. To include a figure either save it as a
#'  ggplot object or indicate a file path to an existing file (possible formats:
#'  png, jpg, bmp).
#'@note For some attributes like plot_widths and plot_heights, if a single value
#'  is provided, it will be reused (behavior of purrr::pmap). This is not the
#'  case for grouplines and group_names. These must be specified for each dataset.
#'@param file file name of the spreadsheet. The extension ".xlsx" is added
#'  automatically.
#'@param datasets A list of an arbitrary number of data.frames, ggplot objects,
#'  and file paths for images in the order in which they should appear in the
#'  output file.
#'@param sheetnames Names of individual worksheets in output file.
#'@param titles Titles shown at the top of the different worksheets.
#'@param sources A list of sources for the different elements of `datasets`.
#'  Elements of this list can also be character vectors to insert more than one source.
#'@param metadata A list containing metadata for each element of `datasets`.
#'  Elements of this list can also be character vectors to insert more than one source.
#'@param grouplines A list containing vectors of indices/names of columns at the beginning of a group.
#'@param group_names A list of character vectors containing the names of the groups
#'  as defined in `grouplines`.
#'@param plot_widths Either a single numeric value denoting the width of all included plots
#'  in inches (1 inch = 2.54 cm), or a list of the same length as `datasets`
#'@param plot_heights Either a single numeric value denoting the height of all included plots
#'  in inches (1 inch = 2.54 cm), or a list of the same length as `datasets`
#'@param index_title Title to be put on the first (overview) sheet.
#'@param index_source Source to be mentioned on the title sheet beneath the title
#'@param logo File path to the logo to be included in the index-sheet. Defaults to the
#'  logo of the Statistical Office of Kanton Zurich.
#'@param contactdetails Character vector with contact information to be displayed
#'  on the title sheet.
#'@param homepage Web address to be put on the title sheet.
#'@param openinghours A character vector with office hours
#'@param auftrag_id An identifier to denote that the output corresponds to a specific
#'  project or order.
#'@param metadata_sheet A list with named elements 'title', 'source', and 'text'.
#'  Intended for conveying long-form information. Default is NULL, not included.
#'@param overwrite Overwrites the existing excel files with the same file name.
#'  default to FALSE
#'@examples
#' library(dplyr)
#' library(statR)
#' library(openxlsx)
#' library(ggplot2)
#'
#'# Example with two datasets and one figure - legacy method
#'fig <- ggplot2::ggplot(mtcars, ggplot2::aes(x = disp))+
#'                  ggplot2::geom_histogram()
#'
#'datasetsXLSX(file = tempfile(fileext = ".xlsx"),
#'             datasets = list(mtcars, PlantGrowth, fig),
#'             titles = c("mtcars-Datensatz",
#'                        "PlantGrowth-Datensatz",
#'                        "Histogramm"),
#'             plot_widths = c(5),
#'             plot_heights = c(5),
#'             sources = list(
#'               paste(
#'                 "Source: Henderson and Velleman (1981).",
#'                 "Building multiple regression models",
#'                 "interactively. Biometrics, 37, 391–411."),
#'               paste(
#'                 "Source: Dobson, A. J. (1983) An Introduction",
#'                 "to Statistical Modelling.",
#'                 "London: Chapman and Hall."),
#'               NULL),
#'             metadata = list(
#'               "Bemerkungen zum mtcars-Datensatz: x",
#'               "Bemerkungen zum PlantGrowth-Datensatz: x",
#'               NULL),
#'             sheetnames = c("Autos","Blumen", "Histogramm"),
#'             index_title = "Autos und Pflanzen",
#'             auftrag_id = "A2021_0000",
#'             overwrite = TRUE)
#'
#' # Newer method
#' df <- mtcars %>%
#'   add_sheetname("cars") %>%
#'   add_title("Cars dataset") %>%
#'   add_source(paste("Henderson and Velleman (1981). Building multiple",
#'                    "regression models interactively.",
#'                    "Biometrics, 37, 391–411."),
#'     prefix = "Source: ", collapse = "") %>%
#'   add_metadata("Obtained in R by calling 'mtcars'",
#'     prefix = "Hinweis: ", collapse = "") %>%
#'   add_grouplines(c(2,5,8)) %>%
#'   add_group_names(c("Group1", "Group2", "Group3"))
#'
#' df2 <- airquality %>%
#'   add_sheetname("airquality") %>%
#'   add_title("Airquality") %>%
#'   add_metadata(c("1. Part of R package 'datasets'",
#'                  "2. Contains some missing values"),
#'     prefix = "Hinweise")
#'
#' plt <- (ggplot(mtcars) + geom_histogram(aes(x = cyl))) %>%
#'   add_sheetname("Histogram") %>%
#'   add_title("A histogram") %>%
#'   add_source("mtcars data from R package 'datasets'",
#'     prefix = "Datenquelle:", collapse = " ") %>%
#'   add_plot_width(6) %>%
#'   add_plot_height(3)
#'
#' datasetsXLSX(
#'   file = "dsxlsx_test.xlsx",
#'   datasets = list(df,  df2, plt),
#'   metadata_sheet = list(
#'     title = "Title of the metadata sheet",
#'     source = "A reference to the responsible organization or similar",
#'     text = c("The metadatasheet is intended for universally applicable",
#'       "long-form explanations which don't fit neatly above the data.",
#'       "",
#'       "Each element is printed in a new row.")
#'   ), overwrite = TRUE)
#'
#' @keywords datasetsXLSX
#' @importFrom dplyr %>%
#' @export
datasetsXLSX <- function(
    file, datasets, sheetnames, titles, sources, metadata, grouplines,
    group_names, plot_widths, plot_heights,
    index_title = getOption("statR_toc_title"),
    index_source = getOption("statR_source"), logo = getOption("statR_logo"),
    contactdetails = inputHelperContactInfo(),
    homepage = getOption("statR_homepage"),
    openinghours = getOption("statR_openinghours"),
    auftrag_id = NULL, metadata_sheet = NULL, overwrite = FALSE) {


  # Try to fill in values if not provided
  if (missing(sheetnames))
    sheetnames <- extract_attributes(datasets, "sheetname", TRUE)

  # Fix: length of sheetnames is truncated to 31 as required by Excel. This is
  # implemented in insert_worksheet_nh. As this function reuses 'sheetnames' to
  # create hyperlinks, the check needs to happen here rather than in the
  # subsequent insert_worksheet_nh calls.
  sheetnames <- verifyInputSheetnames(sheetnames)

  if (missing(titles))
    titles <- extract_attributes(datasets, "title")

  if (missing(sources))
    sources <- extract_attributes(datasets, "source")

  if (missing(metadata))
    metadata <- extract_attributes(datasets, "metadata")

  if (missing(group_names))
    group_names <- extract_attributes(datasets, "group_names")

  if (missing(grouplines))
    grouplines <- extract_attributes(datasets, "grouplines")

  if (all(is.null(metadata_sheet)))
    metadata_sheet <- extract_attribute(datasets, "metadata_sheet")

  # Plot related
  is_plot <- sapply(datasets, checkImplementedPlotType)

  if (any(is_plot)) {

    if (missing(plot_heights)) {
      plot_heights <- extract_attributes(datasets, "plot_height", TRUE)

    } else if (length(plot_heights) == 1) {
      plot_heights <- as.list(ifelse(is_plot, plot_heights, NA))

    } else if (length(plot_heights) == sum(is_plot)) {
      values <- unlist(plot_heights)
      plot_heights <- as.list(rep(NA, length(datasets)))
      plot_heights[which(is_plot)] <- values

    } else if (length(plot_heights) != length(datasets)) {
      stop("Invalid number of values given for plot_heights")
    }

    if (missing(plot_widths)){
      plot_widths <- extract_attributes(datasets, "plot_width", TRUE)

    } else if (length(plot_widths) == 1) {
      plot_widths <- as.list(ifelse(is_plot, plot_widths, NA))

    } else if (length(plot_widths) == sum(is_plot)) {
      values <- unlist(plot_widths)
      plot_widths <- as.list(rep(NA, length(datasets)))
      plot_widths[which(is_plot)] <- values

    } else if (length(plot_widths) != length(datasets)) {
      stop("Invalid number of values given for plot_widths")
    }
  }

  # Initialize new Workbook ------
  wb <- openxlsx::createWorkbook()

  # Insert the initial index sheet ----------
  insert_index_sheet(
    wb = wb, title = index_title, auftrag_id = auftrag_id, logo = logo,
    contactdetails = contactdetails, homepage = homepage,
    openinghours = openinghours, source = index_source)

  # Iterate over datasets
  for (i in seq_along(datasets)) {
    if (checkImplementedPlotType(datasets[[i]])) {
      insert_worksheet_image(
        wb, sheetnames[[i]], image = datasets[[i]], width = plot_widths[[i]],
        height = plot_heights[[i]],
        title = titles[[i]], source = sources[[i]], metadata = metadata[[i]])

    } else if (is.data.frame(datasets[[i]])) {
      insert_worksheet_nh(
        wb = wb, data = datasets[[i]], sheetname = sheetnames[[i]],
        title = titles[[i]], source = sources[[i]], metadata = metadata[[i]],
        grouplines = grouplines[[i]], group_names = group_names[[i]])
    }
  }

  # Create a table of hyperlinks in index sheet (assumed to be "Index") ------
  insert_index_hyperlinks(wb, sheetnames, titles, index_sheet_name = "Index",
                          sheet_start_row = 15)

  # Metadatasheet is intended to receive a list with title, source, and long-form metadata
  # as a character vector, universally applicable and too long to be included with the data.
  if (!is.null(metadata_sheet) && length(metadata_sheet) > 0 && !all(is.na(metadata_sheet))) {
    insert_metadata_sheet(
      wb, sheetname = "Metadatenblatt", title = metadata_sheet[["title"]],
      source = metadata_sheet[["source"]],
      metadata = metadata_sheet[["text"]])
  }

  # Clean unneeded named regions
  cleanNamedRegions(wb, "keep_data")

  # Save workbook at path denoted by argument file ---------
  openxlsx::saveWorkbook(wb, verifyInputFilename(file), overwrite = overwrite)
}


#' splitXLSX()
#'
#' @description Function to export data from R as a formatted .xlsx-file, distributed
#'  over multiple worksheets based on a grouping variable (e.g., year).
#' @note User should make sure that the grouping variable is of binary,
#'   categorical or other types with a limited number of levels.
#' @inheritParams insert_worksheet
#' @param file file name of the output .xlsx-file. The extension is added
#'  automatically.
#' @param sheetvar name of the variable used to split the data and spread them
#'  over several sheets.
#' @examples
#'splitXLSX(file = tempfile(fileext = ".xlsx"),
#'          data = mtcars,
#'          sheetvar = "cyl",
#'          title = "Motor trend car road tests",
#'          source = paste("Source: Henderson and Velleman (1981),",
#'                         "Building multiple regression models interactively.",
#'                         "Biometrics, 37, 391–411."),
#'          metadata = paste("The data was extracted from the 1974 Motor Trend",
#'                           "US magazine and comprises fuel consumption and",
#'                           "10 aspects of automobile design and performance",
#'                           "for 32 automobiles (1973–74 models)."))
#' @keywords splitXLSX
#' @export
splitXLSX <- function(
    file, data, sheetvar, title = "Titel", source = getOption("statR_source"),
    metadata = NA, logo = getOption("statR_logo"),
    contactdetails = inputHelperContactInfo(compact = TRUE),
    homepage = getOption("statR_homepage"),
    author = "user", grouplines = NA, group_names = NA) {

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
#' @note This function is well-suited for applications where a single dataset
#'   needs to be accompanied by a second sheet with explanations or other complex
#'   metadata.
#' @inheritParams insert_worksheet
#' @param file Path of output xlsx-file.
#' @examples
#' aXLSX(file = tempfile(fileext = ".xlsx"),
#'       data = mtcars,
#'       title = "Motor trend car road tests",
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
    file, data, title = "Title", source = getOption("statR_source"),
    metadata = NA, logo = getOption("statR_logo"),
    contactdetails = inputHelperContactInfo(), author = "user", grouplines = NA,
    group_names = NA) {

  # Initialize Workbook object -------
  wb <- openxlsx::createWorkbook()

  # Insert data -----
  insert_worksheet_nh(
    wb, sheetname = "Data", data = data, title = title, source = source,
    metadata = NA, grouplines = grouplines, group_names = group_names)

  # Insert metadata -------
  insert_metadata_sheet(
    wb, title = title, source = source, metadata = metadata, logo = logo,
    contactdetails = contactdetails, author = author)

  # Clean unneeded named regions
  cleanNamedRegions(wb, "keep_data")

  # Write workbook to disk --------
  openxlsx::saveWorkbook(wb, verifyInputFilename(file), overwrite = TRUE)
}


#' quickXLSX()
#'
#' @description A simple function for exporting data from R to a single formatted
#'   .xlsx-spreadsheet.
#' @inheritParams insert_worksheet
#' @param data data to be exported.
#' @param file file name of the xlsx-file. The extension ".xlsx" is added
#' @keywords quickXLSX
#' @export
#' @examples
#'
#' title <- "Motor trend car road tests"
#' source <- paste(
#'   "Henderson and Velleman (1981). Building",
#'   "multiple regression models interactively.",
#'   "Biometrics, 37, 391–411.")
#' metadata <- paste(
#'   "The data was extracted from the 1974 Motor",
#'   "Trend US magazine and comprises fuel consumption",
#'   "and 10 aspects of automobile design and",
#'   "performance for 32 automobiles (1973–74 models).")
#'
#' quickXLSX(file = tempfile(fileext = ".xlsx"),
#'           data = mtcars,
#'           title = title,
#'           source = source,
#'           metadata = metadata)
#'
quickXLSX <- function(
    file, data, title, source,
    metadata, logo = getOption("statR_logo"),
    contactdetails = inputHelperContactInfo(compact = TRUE),
    author = "user", grouplines = NA, group_names = NA) {

  # Create workbook --------
  wb <- openxlsx::createWorkbook()

  # Insert data --------
  insert_worksheet(
    wb, sheetname = "Inhalt", data = data, title = title, source = source,
    metadata = metadata, logo = logo, contactdetails = contactdetails, author = author,
    grouplines = grouplines, group_names = group_names)

  # Clean unneeded named regions
  cleanNamedRegions(wb, "keep_data")

  # Save workbook---------
  openxlsx::saveWorkbook(wb, verifyInputFilename(file), overwrite = TRUE)
}