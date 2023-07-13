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
#'@param index_title Title to be put on the first (overview) sheet.
#'@param datasets datasets or plots to be included.
#'@param plot_widths width of figure in inch (1 inch = 2.54 cm). See details.
#'@param plot_heights height of figure in inch (1 inch = 2.54 cm). See details.
#'@param sheetnames names of the sheet tabs.
#'@param titles titles of the different sheets.
#'@param logo file path to the logo to be included in the index-sheet. Can be
#'  "statzh" or "zh". Defaults to "statzh".
#'@param index_source source to be mentioned on the title sheet beneath the title
#'@param sources source of the data. Defaults to "statzh".
#'@param metadata metadata information to be included. Defaults to NA.
#'@param auftrag_id order number.
#'@param contactdetails contact information on the title sheet. Defaults to "statzh"
#'@param homepage web address to be put on the title sheet. Default to "statzh"
#'@param openinghours openinghours written on the title sheet. Defaults to Data
#'  Shop
#'@param grouplines Column for second header(s). Format: List e.g list(c(2,4,6))
#'@param group_names Name(s) of the second header(s). Format: List e.g
#'  list(c("title 1", "title 2", "title 3"))
#'@param overwrite overwrites the existing excel files with the same file name.
#'  default to FALSE
#'@examples
#'# Example with two datasets and one figure
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
#'             sources = c(paste("Source: Henderson and Velleman (1981).",
#'                               "Building multiple regression models",
#'                               "interactively. Biometrics, 37, 391–411."),
#'                         paste("Source: Dobson, A. J. (1983) An Introduction",
#'                               "to Statistical Modelling.",
#'                               "London: Chapman and Hall.")),
#'             metadata = c("Bemerkungen zum mtcars-Datensatz: x",
#'                          "Bemerkungen zum PlantGrowth-Datensatz: x"),
#'             sheetnames = c("Autos","Blumen", "Histogramm"),
#'             index_title = "Autos und Pflanzen",
#'             auftrag_id = "A2021_0000",
#'             overwrite = TRUE)
#' @keywords datasetsXLSX
#' @importFrom dplyr %>%
#' @export
datasetsXLSX <- function(file,
                         datasets,
                         sheetnames,
                         titles,
                         sources,
                         plot_widths = NULL,
                         plot_heights = NULL,
                         metadata = NA,
                         grouplines = NA,
                         group_names = NA,
                         index_title = getOption("statR_toc_title"),
                         index_source = getOption("statR_source"),
                         logo = getOption("statR_logo"),
                         contactdetails = inputHelperContactInfo(),
                         homepage = getOption("statR_homepage"),
                         openinghours = getOption("statR_openinghours"),
                         auftrag_id = NULL,
                         overwrite = FALSE) {

  # Run checks on arguments ------
  checkGroupOptionCompatibility(group_names, grouplines)

  # Initialize new Workbook ------
  wb <- openxlsx::createWorkbook()

  # Create indexes of which inputs correspond to data.frames or plots-----
  dataframes_index <- which(vapply(datasets, is.data.frame, TRUE))

  implemented_plot_types <- c("gg", "ggplot", "histogram", "character")
  plot_index <- which(vapply(datasets, function(x) {
    length(setdiff(class(x), implemented_plot_types)) == 0
    }, TRUE))

  # Index from input lists using index -----------
  ### data.frames
  dataframe_datasets <- datasets[dataframes_index]
  dataframe_sheetnames <- sheetnames[dataframes_index]
  dataframe_titles <- titles[dataframes_index]
  dataframe_sources <- sources[dataframes_index]
  dataframe_metadata <- metadata[dataframes_index]
  dataframe_grouplines <- grouplines[dataframes_index]
  dataframe_group_names <- group_names[dataframes_index]

  ### Plots
  plot_datasets <- datasets[plot_index]
  plot_sheetnames <- sheetnames[plot_index]


  # Insert the initial index sheet ----------
  insert_index_sheet(wb = wb,
                     title = index_title,
                     auftrag_id = auftrag_id,
                     logo = logo,
                     contactdetails = contactdetails,
                     homepage = homepage,
                     openinghours = openinghours,
                     source = index_source)


  # Insert datasets according to dataframes_index -------
  if (length(dataframes_index) > 0) {
    list(dataframe_datasets,
         dataframe_sheetnames,
         dataframe_titles,
         dataframe_sources,
         dataframe_metadata,
         dataframe_grouplines,
         dataframe_group_names) %>%
      purrr::pwalk(~insert_worksheet_nh(wb = wb,
                                        data = ..1,
                                        sheetname = ..2,
                                        title = ..3,
                                        source = ..4,
                                        metadata = ..5,
                                        grouplines = ..6,
                                        group_names = ..7))
  }


  # Insert images according to plot_index --------
  if (length(plot_index) > 0) {
    list(plot_datasets,
         plot_sheetnames,
         plot_widths,
         plot_heights) %>%
      purrr::pmap(~insert_worksheet_image(wb = wb,
                                          image = ..1,
                                          sheetname = ..2,
                                          width = ..3,
                                          height = ..4))
  }


  # Create a table of hyperlinks in index sheet (assumed to be "Index") ------
  insert_hyperlinks(wb, sheetnames, titles, index_sheet_name = "Index",
                    sheet_start_row = 15)


  # Save workbook at path denoted by argument file ---------
  openxlsx::saveWorkbook(wb, verifyInputFilename(file), overwrite = overwrite)
}
