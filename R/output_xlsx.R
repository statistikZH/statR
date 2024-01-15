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
#' @param file file name of the spreadsheet. The extension ".xlsx" is added
#'  automatically.
#'
#' @param datasets A list of an arbitrary number of data.frames, ggplot objects,
#'  and file paths for images in the order in which they should appear in the
#'  output file.
#'
#' @param sheetnames Names of individual worksheets in output file. Note that these
#'   will be truncated to 31 characters and must be unique.
#'
#' @param titles Titles shown at the top of the different worksheets.
#'
#' @param sources A list of sources for the different elements of `datasets`.
#'  Elements of this list can also be character vectors to insert more than one source.
#'
#' @param metadata A list containing metadata for each element of `datasets`.
#'  Elements of this list can also be character vectors to insert more than one source.
#'
#' @param grouplines A list containing vectors of indices/names of columns at
#'   the beginning of a group.
#'
#' @param group_names A list of character vectors containing the names of the
#'   groups as defined in `grouplines`. Should not be specified unless
#'   grouplines is also specified.
#'
#' @param plot_widths Either a single numeric value denoting the width of all included plots
#'  in inches (1 inch = 2.54 cm), or a list of the same length as `datasets`
#'
#' @param plot_heights Either a single numeric value denoting the height of all included plots
#'  in inches (1 inch = 2.54 cm), or a list of the same length as `datasets`
#'
#' @param index_title Title to be put on the index sheet.
#'
#' @param index_source Source to be shown below the index title.
#'
#' @param logo File path to the logo to be included in the index-sheet. Defaults to the
#'  logo of the Statistical Office of Kanton Zurich.
#'
#' @param contactdetails Character vector with contact information to be displayed
#'  on the title sheet. By default uses \code{inputHelperContactInfo()} to
#'  construct it based on the user config user defined values.
#'
#' @param homepage Web address to be put on the title sheet.
#'
#' @param openinghours A character vector with office hours
#'
#' @param auftrag_id An identifier to denote that the output corresponds to a specific
#'  project or order.
#'
#' @param metadata_sheet A list with named elements 'title', 'source', and 'text'.
#'  Intended for conveying long-form information. Default is NULL, not included.
#' @param overwrite Overwrites the existing excel files with the same file name.
#'  default to TRUE
#' @examples
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
    file, datasets, sheetnames = NULL, titles = NULL, sources = NULL, metadata = NULL, grouplines = NULL,
    group_names = NULL, plot_widths = NULL, plot_heights = NULL,
    index_title = getOption("statR_toc_title"),
    index_source = getOption("statR_source"), logo = getOption("statR_logo"),
    contactdetails = inputHelperContactInfo(),
    homepage = getOption("statR_homepage"),
    openinghours = getOption("statR_openinghours"),
    auftrag_id = NULL, metadata_sheet = NULL, overwrite = TRUE) {




  if(!is.null(titles)){
    datasets <- purrr::map2(datasets, titles, ~add_title(..1, ..2))
  }
  if(!is.null(sources)){
    datasets <- purrr::map2(datasets, sources, ~add_source(..1, ..2))
  }
  if(!is.null(metadata)){
    datasets <- purrr::map2(datasets, metadata, ~add_metadata(..1, ..2))
  }
  if(!is.null(grouplines)){
    datasets <- purrr::map2(datasets, grouplines, ~add_grouplines(..1, ..2))
  }
  if(!is.null(group_names)){
    datasets <- purrr::map2(datasets, group_names, ~add_group_names(..1, ..2))
  }

  if(is.null(sheetnames)){
    sheetnames <- extract_attributes(datasets, "sheetname")
  }
  if(is.null(plot_heights)){
    plot_heights <- extract_attributes(datasets, "plot_height")
    plot_widths <- extract_attributes(datasets, "plot_width")
  }


  # Fix: length of sheetnames is truncated to 31 as required by Excel. This is
  # implemented in insert_worksheet_nh, but this function reuses 'sheetnames' to
  # create hyperlinks. Hence the check needs to happen here rather than in the
  # subsequent insert_worksheet_nh calls to avoid dead links.
  sheetnames <- verifyInputSheetnames(sheetnames)

  # Plot related
  is_plot <- sapply(datasets, checkImplementedPlotType)

  if (any(is_plot)) {

    if (length(plot_heights) == 1) {
      plot_heights <- as.list(ifelse(is_plot, plot_heights, NA))

    } else if (length(plot_heights) == sum(is_plot)) {
      values <- unlist(plot_heights)
      plot_heights <- as.list(rep(NA, length(datasets)))
      plot_heights[which(is_plot)] <- values

    } else if (length(plot_heights) != length(datasets)) {
      stop("Invalid number of values given for plot_heights")
    }


    if (length(plot_widths) == 1) {
      plot_widths <- as.list(ifelse(is_plot, plot_widths, NA))

    } else if (length(plot_widths) == sum(is_plot)) {
      values <- unlist(plot_widths)
      plot_widths <- as.list(rep(NA, length(datasets)))
      plot_widths[which(is_plot)] <- values

    } else if (length(plot_widths) != length(datasets)) {
      stop("Invalid number of values given for plot_widths")
    }

    plot_size_list <- purrr::map2(plot_heights, plot_widths, ~c(..1,..2))

    datasets <- purrr::map2(datasets, plot_size_list, ~add_plot_size(..1, ..2))
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
      insert_worksheet_image(wb, sheetname = sheetnames[i], image = datasets[[i]])

    } else if (is.data.frame(datasets[[i]])) {
      insert_worksheet_nh(wb, sheetname = sheetnames[i], data = datasets[[i]])
    }
  }





  # Create a table of hyperlinks in index sheet (assumed to be "Index") ------
  titles <- extract_attributes(datasets,"title")

  insert_index_hyperlinks(wb, sheetnames, titles = titles, index_sheet_name = "Index",
                          sheet_start_row = namedRegionLastRow(wb, "Index", "toc") + 1)

  # Metadata sheets are constructed from a list with title, source, and
  # long-form metadata by insert_metadata_sheet. It is meant to be used to
  # provide globally applicable information for multiple analyses.
  if (!is.null(metadata_sheet) && length(metadata_sheet) > 0 && !all(is.na(metadata_sheet))) {
    insert_metadata_sheet(
      wb, sheetname = "Metadatenblatt", meta_infos = metadata_sheet)
  }

  # Clean unneeded named regions - keep_data or all. Using 'all' here as a
  # quick fix for region names with non-standard characters. See NEWS.md
  cleanNamedRegions(wb, "all")

  # Save workbook at path denoted by argument file ---------
  openxlsx::saveWorkbook(wb, verifyInputFilename(file), overwrite = overwrite)
}


#' Export a dataset split by a covariate into a workbook
#'
#' @description Function to export data from R as a formatted .xlsx-file,
#'   distributed over multiple worksheets based on a grouping variable (e.g., year).
#' @note User should make sure that the grouping variable is of binary,
#'   categorical or other types with a limited number of levels.
#' @inheritParams quickXLSX
#' @param homepage Web address to be put on the title sheet.
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
    file, data, sheetvar, title, source, metadata, grouplines = NA,
    group_names = NA, logo = getOption("statR_logo"),
    contactdetails = inputHelperContactInfo(compact = TRUE),
    homepage = getOption("statR_homepage"),
    author = "user") {



  datasets <- split.data.frame(data, data[,sheetvar])
  sheetnames <- paste0(sheetvar, "_", names(datasets))
  titles <- paste0(title, " (", sheetvar, ": ", names(datasets), ")")

  datasets <- purrr::map2(datasets, titles, ~add_title(..1, ..2))
  datasets <- lapply(datasets, function(x) add_metadata(x, metadata))
  datasets <- lapply(datasets, function(x) add_source(x, source))
  datasets <- lapply(datasets, function(x) add_grouplines(x, grouplines))
  datasets <- lapply(datasets, function(x) add_group_names(x, group_names))


  datasetsXLSX(file = file, datasets = datasets, sheetnames = sheetnames)
}





#' Export a single dataset to a workbook with an index sheet
#'
#' @description Function to export data from R to a formatted .xlsx-file. The
#'  data is exported to the first sheet. Metadata information is exported to
#'  the second sheet.
#' @note This function is well-suited for applications where a single dataset
#'   needs to be accompanied by a second sheet with explanations or other complex
#'   metadata.
#' @inheritParams quickXLSX
#' @param meta_infos metadata information as a list
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
    file, data, title = NULL, source = NULL, metadata = NULL, grouplines = NULL, group_names = NULL,
    logo = getOption("statR_logo"), contactdetails = inputHelperContactInfo(),
    author = "user") {


  if(!is.null(title)){
    data <- add_title(data, title)
  }
  if(!is.null(source)){
    data <- add_source(data, source)
  }
  if(!is.null(metadata)){
    data <- add_metadata(data, metadata)
  }
  if(!is.null(grouplines)){
    data <- add_grouplines(data, grouplines)
  }
  if(!is.null(group_names)){
    data <- add_group_names(data, group_names)
  }

  class(data) <- c(class(data), "Content")



  wb <- openxlsx::createWorkbook()

  # Insert data -----
  insert_worksheet_nh(wb, sheetname = "Data", data = data, metadata = NA)


  meta_info <- list(
    title = extract_attribute(data, "title"),
    source = extract_attribute(data, "source"),
    metadata = extract_attribute(data, "metadata")
  )

  # Insert metadata -------
  insert_metadata_sheet(wb, sheetname = "Metadaten", meta_infos = meta_info, logo = logo, contactdetails = contactdetails, author = author)

  # Clean unneeded named regions
  cleanNamedRegions(wb, "all")

  # Write workbook to disk --------
  openxlsx::saveWorkbook(wb, verifyInputFilename(file), overwrite = TRUE)
}





#' Export a single dataset to a single formatted worksheet
#'
#' @description A simple function for exporting data from R to a single formatted
#'   .xlsx-spreadsheet.
#' @inheritParams insert_worksheet
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
    file, data, title = NULL, source = NULL, metadata = NULL, grouplines = NULL, group_names = NULL,
    logo = getOption("statR_logo"),
    contactdetails = statR:::inputHelperContactInfo(compact = TRUE),
    author = "user") {


  if(!is.null(title)){
    data <- add_title(data, title)
  }
  if(!is.null(source)){
    data <- add_source(data, source)
  }
  if(!is.null(metadata)){
    data <- add_metadata(data, metadata)
  }
  if(!is.null(grouplines)){
    data <- add_grouplines(data, grouplines)
  }
  if(!is.null(group_names)){
    data <- add_group_names(data, group_names)
  }

  class(data) <- c(class(data), "Content")


  # Create workbook --------
  wb <- openxlsx::createWorkbook()

  # Insert data --------
  insert_worksheet(wb, sheetname = "Inhalt", data = data)

  # Clean unneeded named regions
  cleanNamedRegions(wb, "all")

  # Save workbook---------
  openxlsx::saveWorkbook(wb, verifyInputFilename(file), overwrite = TRUE)
}
