#' datasetsXLSX()
#'
#' Function to export several datasets and/or figures from R to an .xlsx-file. The function creates an overview sheet and separate sheets
#' for each dataset/figure.
#'
#' When including figures, the heights and widths need to be specified as a vector. For example, say you have one dataset and two figures
#' that you would like to export. widths = c(5,6) then suggests that the first figure will be 5 inches wide, the second 6. To include a figure either
#' save it as a ggplot object or indicate a file path to an existing file (possible formats: png, jpg, bmp).
#'
#' @param file file name of the spreadsheet. The extension ".xlsx" is added automatically.
#' @param maintitle Title to be put on the first (overview) sheet.
#' @param datasets datasets or plots to be included.
#' @param plot_widths width of figure in inch (1 inch = 2.54 cm). See details.
#' @param plot_heights height of figure in inch (1 inch = 2.54 cm). See details.
#' @param sheetnames names of the sheet tabs.
#' @param titles titles of the different sheets.
#' @param logo file path to the logo to be included in the index-sheet. Can be "statzh" or "zh". Defaults to "statzh".
#' @param titlesource source to be mentioned on the title sheet beneath the title
#' @param sources source of the data. Defaults to "statzh".
#' @param metadata1 metadata information to be included. Defaults to NA.
#' @param auftrag_id order number.
#' @param contact contact information on the title sheet. Defaults to "statzh"
#' @param homepage web address to be put on the title sheet. Default to "statzh"
#' @param openinghours openinghours written on the title sheet. Defaults to Data Shop
#' @param grouplines Column for second header(s). Format: List e.g list(c(2,4,6))
#' @param group_names Name(s) of the second header(s). Format: List e.g list(c("title 1", "title 2", "title 3"))
#' @param overwrite overwrites the existing excel files with the same file name. default to FALSE
#' @keywords datasetsXLSX
#' @export
#' @importFrom dplyr "%>%"
#' @importFrom purrr pwalk pmap
#' @examples
#'\donttest{
#' \dontrun{
#'# Example with two datasets and no figure
#'dat1 <- mtcars
#'dat2 <- PlantGrowth
#'
#'datasetsXLSX(file="twoDatasets",
#'             datasets = list(dat1, dat2),
#'             titles = c("mtcars-Datensatz","PlantGrowth-Datensatz"),
#'             grouplines = list(c(1)),
#'             # adds a second header in the first sheet
#'             group_names = list(c("name_of_second_header")),
#'             sources = c("Source: Henderson and Velleman (1981).
#'             Building multiple regression models interactively. Biometrics, 37, 391–411.",
#'                         "Dobson, A. J. (1983) An Introduction to Statistical
#'                         Modelling. London: Chapman and Hall."),
#'             metadata1 = c("Bemerkungen zum mtcars-Datensatz: x",
#'                           "Bemerkungen zum PlantGrowth-Datensatz: x"),
#'             sheetnames = c("Autos","Blumen"),
#'             maintitle = "Autos und Pflanzen",
#'             titlesource = "statzh",
#'             logo = "statzh",
#'             auftrag_id="A2021_0000",
#'             contact = "statzh",
#'             homepage = "statzh",
#'             openinghours = "statzh",
#'             overwrite = T)
#'
#'# Example with two datasets and one figure
#'
#'dat1 <- mtcars
#'dat2 <- PlantGrowth
#'fig <- ggplot(mtcars, aes(x=disp))+
#'                  geom_histogram()
#'
#'datasetsXLSX(file="twoDatasetsandFigure",
#'             datasets = list(dat1, dat2, fig),   # fig als ggplot Objekt oder File Path
#'             titles = c("mtcars-Datensatz","PlantGrowth-Datensatz", "Histogramm"),
#'             plot_widths = c(5),
#'             plot_heights = c(5),
#'             sources = c("Source: Henderson and Velleman (1981).
#'             Building multiple regression models interactively. Biometrics, 37, 391–411.",
#'                         "Source: Dobson, A. J. (1983) An Introduction to
#'                         Statistical Modelling. London: Chapman and Hall."),
#'             metadata1 = c("Bemerkungen zum mtcars-Datensatz: x",
#'                          "Bemerkungen zum PlantGrowth-Datensatz: x"),
#'             sheetnames = c("Autos","Blumen", "Histogramm"),
#'             maintitle = "Autos und Pflanzen",
#'             titlesource = "statzh",
#'             logo="statzh",
#'             auftrag_id="A2021_0000",
#'             contact = "statzh",
#'             homepage = "statzh",
#'             openinghours = "statzh",
#'             overwrite = T)
#'}
#'}

datasetsXLSX <- function(file, datasets, titles, plot_widths = NULL,
                         plot_heights = NULL, grouplines = NA,
                         group_names = NA, sources = "statzh", metadata1 = NA,
                         sheetnames, maintitle, titlesource = "statzh",
                         logo = "statzh", auftrag_id = NULL,
                         contact = "statzh", homepage = "statzh",
                         openinghours = "statzh", overwrite = F){

  if(!any(is.na(group_names)) & any(is.na(grouplines))){
    stop("if a second header is wanted, the grouplines have to be specified")
  }

  wb <- openxlsx::createWorkbook()

  # Determine which elements of input list datasets correspond to dataframes.
  dataframes_index <- which(vapply(datasets, is.data.frame, TRUE))

  dataframe_datasets <- datasets[dataframes_index]
  dataframe_sheetnames <- sheetnames[dataframes_index]
  dataframe_titles <- titles[dataframes_index]
  dataframe_sources <- sources[dataframes_index]
  dataframe_metadata1 <- metadata1[dataframes_index]
  dataframe_grouplines <- grouplines[dataframes_index]
  dataframe_group_names <- group_names[dataframes_index]


  # Determine which elements of datasets correspond to objects of type gg,
  # ggplot, histogram, or character (path input).
  plot_index <- which(vapply(datasets, function(x){
    length(setdiff(class(x), c("gg", "ggplot", "histogram", "character"))) == 0
    }, TRUE))

  plot_datasets <- datasets[plot_index]
  plot_sheetnames <- sheetnames[plot_index]


  # Insert the initial index sheet
  insert_index_sheet(wb, logo, contact, homepage, openinghours, titlesource,
                     auftrag_id, maintitle)


  # Iterate along dataframes_index
  if(length(dataframes_index) > 0){
    list(
      dataframe_datasets,
      dataframe_sheetnames,
      dataframe_titles,
      dataframe_sources,
      dataframe_metadata1,
      dataframe_grouplines,
      dataframe_group_names
      ) %>%
      purrr::pwalk(
        ~insert_worksheet_nh(
          data = ..1,
          wb = wb,
          sheetname = ..2,
          title = ..3,
          source = ..4,
          metadata = ..5,
          grouplines = ..6,
          group_names = ..7))
  }

  # Iterate along plot_index
  if (length(plot_index) > 0){
    list(
      plot_datasets,
      plot_sheetnames,
      plot_widths,
      plot_heights
    ) %>%
      purrr::pmap(
        ~insert_worksheet_image(
          image = ..1,
          wb = wb,
          sheetname = ..2,
          width = ..3,
          height = ..4))
  }

  # Create a table of hyperlinks
  data.frame(
    sheetnames = sheetnames,
    titles = titles,
    sheet_row = c(seq(15,15+length(sheetnames)-1))
  ) %>%
  purrr::pwalk(
    ~insert_hyperlinks(
      wb,
      sheetname = ..1,
      title = ..2,
      sheet_row = ..3))

  # Save workbook at path denoted by argument file
  openxlsx::saveWorkbook(wb, file = prep_filename(file), overwrite = overwrite)
}
