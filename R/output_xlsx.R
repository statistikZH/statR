#' Export datasets and graphics into a multi-worksheet workbook
#'
#' Function to export multiple datasets and/or figures from R to a workbook.
#' The function creates an index sheet, separate worksheets for each input
#' object, hyperlinks to each content worksheet, and an optional sheet for
#' long-form metadata.
#'
#' @details The arguments datasets, sheetnames, titles, sources, and metadata
#'  should be of equal length. For more complex workbooks, it may be more
#'  convenient to parametrize these for each individual input using functions
#'  like \code{add_sheetname()}, which can be chained using %>% operators.
#'
#'  All elements in datasets must be of type data.frame, ggplot, or character.
#'  In the latter case, the inputs must correspond to paths to existing image
#'  files.
#' @inheritParams insert_worksheet
#' @inheritParams quickXLSX
#' @param datasets A list of an arbitrary number of data.frames, ggplot objects,
#'  and file paths for images in the order in which they should appear in the
#'  output file.
#' @param sheetname Names of individual worksheets in output file. Note that
#'   these names will be truncated to 31 characters, must be unique, and cannot
#'   contain some special characters (namely the following: /, \, ?, *, :, [, ]).
#' @param title Titles shown at the top of the different worksheets.
#' @param source A list of sources for the different elements of `datasets`.
#'  Elements of this list can also be character vectors to insert more than one source.
#' @param metadata A list containing metadata for each element of `datasets`.
#'  Elements of this list can also be character vectors to insert more than one source.
#' @param grouplines A list containing vectors of indices/names of columns at
#'   the beginning of a group.
#' @param group_names A list of character vectors containing the names of the
#'   groups as defined in `grouplines`. Should not be specified unless
#'   grouplines is also specified.
#' @param plot_width Either a single numeric value denoting the width of all included plots
#'  in inches (1 inch = 2.54 cm), or a list of the same length as `datasets`
#' @param plot_height Either a single numeric value denoting the height of all included plots
#'  in inches (1 inch = 2.54 cm), or a list of the same length as `datasets`
#' @param index_title Title to be put on the index sheet.
#' @param index_source Source to be shown below the index title.
#' @param metadata_sheet A list with named elements 'title', 'source', and 'text'.
#'  Intended for conveying long-form information. Default is NULL, not included.
#' @param overwrite Overwrites the existing excel files with the same file name.
#'  default to TRUE
#' @param config which config file should be used. Default: default
#' @examples
#' library(dplyr)
#' library(statR)
#' library(ggplot2)
#'
#'# Example with two datasets and one figure - legacy method
#' fig <- ggplot2::ggplot(mtcars, ggplot2::aes(x = disp)) +
#'   ggplot2::geom_histogram()
#'
#' \dontrun{
#' datasetsXLSX(file = tempfile(fileext = ".xlsx"),
#'              datasets = list(mtcars, PlantGrowth, fig),
#'              title = c("mtcars-Datensatz",
#'                        "PlantGrowth-Datensatz",
#'                        "Histogramm"),
#'              plot_width = c(5),
#'              plot_height = c(5),
#'              source = list(
#'                paste(
#'                  "Source: Henderson and Velleman (1981).",
#'                  "Building multiple regression models",
#'                  "interactively. Biometrics, 37, 391–411."),
#'                paste(
#'                  "Source: Dobson, A. J. (1983) An Introduction",
#'                  "to Statistical Modelling.",
#'                  "London: Chapman and Hall."),
#'                NULL),
#'              metadata = list(
#'                "Bemerkungen zum mtcars-Datensatz: x",
#'                "Bemerkungen zum PlantGrowth-Datensatz: x",
#'                NULL),
#'              sheetname = c("Autos","Blumen", "Histogramm"),
#'              index_title = "Autos und Pflanzen",
#'              auftrag_id = "A2021_0000",
#'              overwrite = TRUE)
#' }
#'
#' # Newer method
#' df <- mtcars %>%
#'   add_sheetname("cars") %>%
#'   add_title("Cars dataset") %>%
#'   add_source(paste("Henderson and Velleman (1981). Building multiple",
#'                    "regression models interactively.",
#'                    "Biometrics, 37, 391–411.")) %>%
#'   add_metadata("Obtained in R by calling 'mtcars'") %>%
#'   add_grouplines(c(2, 5, 8)) %>%
#'   add_group_names(c("Group1", "Group2", "Group3"))
#'
#' df2 <- airquality %>%
#'   add_sheetname("airquality") %>%
#'   add_title("Airquality") %>%
#'   add_metadata("Obtained in R by calling 'airquality'")
#'
#' plt <- (ggplot(mtcars) + geom_histogram(aes(x = cyl))) %>%
#'   add_sheetname("Histogram") %>%
#'   add_title("A histogram") %>%
#'   add_source("mtcars data from R package 'datasets'") %>%
#'   add_plot_size(c(6,3))
#'
#' metadata_sheet <- list(
#'   title = "Title of the metadata sheet",
#'   source = "A reference to the responsible organization or similar",
#'   text = c("The metadatasheet is intended for universally applicable",
#'            "long-form explanations which don't fit neatly above the data.",
#'            "",
#'            "Each element is printed in a new row."))
#'
#' \dontrun{
#' datasetsXLSX(
#'   file = tempfile(fileext = ".xlsx"),
#'   datasets = list(df, df2, plt),
#'   metadata_sheet = metadata_sheet,
#'   overwrite = TRUE)
#' }
#'
#' @keywords datasetsXLSX
#' @importFrom dplyr %>%
#' @importFrom stats na.omit
#' @export
datasetsXLSX <- function(
    file, datasets, sheetname = NULL, title = NULL, source = NULL,
    metadata = NULL, grouplines = NULL, group_names = NULL, plot_width = NULL,
    plot_height = NULL, index_title = NA,
    index_source = NA, logo = NA,
    contactdetails = NA,
    homepage = NA,
    openinghours = NA,
    auftrag_id = NA, author = "user", metadata_sheet = NULL, overwrite = TRUE, config = "default") {


  user_config <- get_user_config(config, c(index_title, index_source, logo, contactdetails, homepage, openinghours))

  # "Optional" input arguments
  for (value in c("title", "source", "metadata", "grouplines", "group_names")){
    if (!is.null(eval(as.name(value)))) {
      datasets <- purrr::map2(datasets, eval(as.name(value)),
                              ~add_attribute(..1, value, ..2))
    }
  }

  # Required input arguments
  for (value in c("sheetname")){
    if (is.null(eval(as.name(value)))) {
      assign(value, extract_attributes(datasets, value, required_val = TRUE))
    }
  }

  # Plot related arguments
  is_plot <- sapply(datasets, checkImplementedPlotType)

  if (any(is_plot)) {

    # Reads object with name == arg, if NULL tries to extract attribute.
    # If length == 1, all plots are assigned the same plot size, otherwise
    # assigns plot_size in order of plots. Overrides all attributes if non-null
    # plot_size is provided.
    for (arg in c("plot_height", "plot_width")){
      buffer <- as.list(rep(NA, length.out = length(datasets)))

      if (is.null((values<- eval(as.name(arg))))) {
        values <- unlist(extract_attributes(datasets, arg))
      }

      if (!(sum(!is.na(values)) %in% c(1, sum(is_plot)))) {
        stop("Invalid number of values for ", arg)
      }

      buffer[which(is_plot)] <- na.omit(values)
      assign(arg, buffer)
    }

    plot_size_list <- purrr::map2(plot_width, plot_height, ~c(..1,..2))
    datasets <- purrr::map2(datasets, plot_size_list, ~add_plot_size(..1, ..2))
  }

  sheetname <- verifyInputSheetnames(sheetname)

  wb <- openxlsx::createWorkbook()
  insert_index_sheet(wb, sheetname = "Index", title = index_title,
                     auftrag_id = auftrag_id, logo = logo,
                     contactdetails = contactdetails, homepage = homepage,
                     openinghours = openinghours, source = index_source,
                     author = author)

  # Iterate over datasets
  for (i in seq_along(datasets)) {
    if (is_plot[i]) {
      insert_worksheet_image(wb, sheetname = sheetname[i], image = datasets[[i]])

    } else if (is.data.frame(datasets[[i]])) {
      insert_worksheet_nh(wb, sheetname = sheetname[i], data = datasets[[i]])
    }
  }

  insert_index_hyperlinks(wb, sheetname, extract_attributes(datasets, "title"),
                          index_sheet_name = "Index",
                          sheet_start_row = namedRegionLastRow(wb, "Index", "toc") + 1)

  # Metadata sheets are constructed from a list with title, source, and
  # long-form metadata by insert_metadata_sheet. It is meant to be used to
  # provide globally applicable information for multiple analyses.
  if (!is.null(metadata_sheet)) {
    insert_metadata_sheet(wb, "Beiblatt", metadata_sheet, logo = logo,
                          contactdetails = contactdetails, homepage = homepage,
                          author = author)
  }

  # Clean unneeded named regions - keep_data or all. Using 'all' here as a
  # quick fix for region names with non-standard characters. See NEWS.md
  cleanNamedRegions(wb, "all")
  openxlsx::saveWorkbook(wb, verifyInputFilename(file), overwrite = overwrite)
}


#' Export a dataset split by a covariate into a workbook
#'
#' @description Function to export data from R as a formatted .xlsx-file,
#'   distributed over multiple worksheets based on a grouping variable (e.g., year).
#' @note User should make sure that the grouping variable is of binary,
#'   categorical or other types with a limited number of levels.
#' @inheritParams insert_worksheet
#' @inheritParams quickXLSX
#' @param sheetvar name of the variable used to split the data and spread them
#'  over several sheets.
#' @examples
#' \dontrun{
#' splitXLSX(file = tempfile(fileext = ".xlsx"),
#'           data = mtcars,
#'           sheetvar = "cyl",
#'           title = "Motor trend car road tests",
#'           source = paste("Source: Henderson and Velleman (1981),",
#'                          "Building multiple regression models interactively.",
#'                          "Biometrics, 37, 391–411."),
#'           metadata = paste("The data was extracted from the 1974 Motor Trend",
#'                            "US magazine and comprises fuel consumption and",
#'                            "10 aspects of automobile design and performance",
#'                            "for 32 automobiles (1973–74 models)."))
#' }
#' @keywords splitXLSX
#' @export
splitXLSX <- function(
    file, data, sheetvar, title = NULL, source = NULL, metadata = NULL,
    grouplines = NULL, group_names = NULL, logo = NA,
    contactdetails = NA,
    homepage = NA, author = "user", config = "default") {

  get_user_config(config, c(logo, contactdetails,  homepage))


  # Shared values: these are attached to the source data.frame before
  # splitting on sheetvar using split.data.frame(), which preserves attributes
  for (value in c("source", "metadata", "grouplines", "group_names")) {
    if (!is.null(eval(as.name(value)))) {
      data <- add_attribute(data, value, eval(as.name(value)))
    }
  }

  datasets <- split.data.frame(data, data[,sheetvar])
  sheetnames <- paste0(sheetvar, "_", names(datasets))
  titles <- paste0(title, " (", sheetvar, ": ", names(datasets), ")")

  datasetsXLSX(file = file, datasets = datasets, sheetname = sheetnames,
               title = titles, logo = logo, contactdetails = contactdetails,
               homepage = homepage, author = author)
}


#' Export a single dataset to a workbook with an index sheet
#'
#' @description Function to export data from R to a formatted .xlsx-file. The
#'  data is exported to the first sheet. Metadata information is exported to
#'  the second sheet.
#' @note This function is well-suited for applications where a single dataset
#'   needs to be accompanied by a second sheet with explanations or other complex
#'   metadata.
#' @inheritParams insert_worksheet
#' @inheritParams quickXLSX
#' @examples
#' \dontrun{
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
#' }
#' @keywords aXLSX
#' @export
aXLSX <- function(
    file, data, title = NULL, source = NULL, metadata = NULL, grouplines = NULL,
    group_names = NULL, logo = NA,
    contactdetails = NA,
    homepage = NA, author = "user", config = "default") {

  get_user_config(config, c(logo, contactdetails,  homepage))


  for (value in c("title", "source", "metadata", "grouplines", "group_names")) {
    if (!is.null(eval(as.name(value)))) {
      data <- add_attribute(data, value, eval(as.name(value)))
    }
  }

  meta_info_list <- list(
    title = extract_attribute(data, "title"),
    source = extract_attribute(data, "source"),
    metadata = extract_attribute(data, "metadata")
  )

  wb <- openxlsx::createWorkbook()
  insert_worksheet_nh(wb, sheetname = "Data", data = data, metadata = NA)
  insert_metadata_sheet(wb, "Metadaten", meta_info_list, logo,
                        contactdetails, homepage, author)
  cleanNamedRegions(wb, "all")
  openxlsx::saveWorkbook(wb, verifyInputFilename(file), overwrite = TRUE)
}


#' Export a single dataset to a single formatted worksheet
#'
#' @description A simple function for exporting data from R to a single formatted
#'   .xlsx-spreadsheet.
#' @inheritParams insert_worksheet
#' @param config which config file should be used. Default: default
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
#' \dontrun{
#' quickXLSX(file = tempfile(fileext = ".xlsx"),
#'           data = mtcars,
#'           title = title,
#'           source = source,
#'           metadata = metadata)
#' }
quickXLSX <- function(
    data, file, title = NULL, source = NULL, metadata = NULL, grouplines = NULL,
    group_names = NULL, logo = NA,
    contactdetails = NA,
    homepage = NA,
    author = "user",
    config = "default") {


  get_user_config(config, list(logo, contactdetails,  homepage))



  for (value in c("title", "source", "metadata", "grouplines", "group_names")) {
    if (!is.null(eval(as.name(value)))) {
      data <- add_attribute(data, value, eval(as.name(value)))
    }
  }

  wb <- openxlsx::createWorkbook()
  insert_worksheet(wb, sheetname = "Inhalt", data = data, logo = logo,
                   contactdetails = getOption("statR_contactdetails"), homepage = homepage,
                   author = author)
  cleanNamedRegions(wb, "all")
  openxlsx::saveWorkbook(wb, verifyInputFilename(file), overwrite = TRUE)
}
