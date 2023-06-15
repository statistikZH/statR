#' insert_worksheet_image()
#'
#' @description Insert an image into a worksheet
#' @note The function does not write the result into a .xlsx file. A separate
#'   call to openxlsx::saveWorkbook() is required. A temporary file is created
#'   for inputs of type gg, ggplot or histogram object at path given by
#'   `tempfile()`.
#' @param image image or plot
#' @param wb workbook object to write new worksheet in
#' @param sheetname  name of the sheet tab
#' @param startrow row coordinate of upper left corner of figure
#' @param startcol column coordinate of upper left corner of figure
#' @param width width of figure
#' @param height height of figure
#' @param units unit of measurement (default: in)
#' @param dpi image resolution (default: 300)
#' @examples
#' figure <- ggplot2::ggplot(mtcars, ggplot2::aes(x = disp)) +
#'   ggplot2::geom_histogram()
#'
#' wb <- openxlsx::createWorkbook()
#' insert_worksheet_image(wb, sheetname = "ggplot image",
#'                        image = figure, width = 3.5, height = 5.5)
#' @keywords insert_worksheet_image
#' @importFrom methods is
#' @export
insert_worksheet_image <- function(wb, sheetname, image, startrow = 3,
                                   startcol = 3, width, height,
                                   units = "in", dpi = 300){

  # If image is.null, pass
  if (is.null(image)) return()

  else if (is.character(image)) image_path <- image

  else if (checkImplementedPlotType(image)){
    image_path <- tempfile(fileext = ".png")
  }

  else {
    stop("Plot muss als ggplot Objekt oder als Filepath vorliegen.")
  }


  # Create png of plot ----------
  if (is(image, "histogram")){
    grDevices::png(image_path, width = width, height = height, units = units,
                   res = dpi)
    plot(image)
    grDevices::dev.off()
  }

  else if (is(image, "gg") | is(image, "ggplot2")){
    ggplot2::ggsave(image_path, plot = image, device = "png",
                    width = width, height = height, units = units, dpi = dpi)
  }


  # If file not found at image_path, warn and pass
  if (!file.exists(image_path)){
    warning("Image not found.")
    return()
  }


  # Add worksheet if not already defined ---------
  if (!(sheetname %in% names(wb))){
    openxlsx::addWorksheet(wb, sheetName = sheetname, gridLines = FALSE)
  }

  # Insert image ---------
  openxlsx::insertImage(wb, sheet = sheetname, file = image_path,
                        width = width, height = height, startRow = startrow,
                        startCol = startcol, units = units, dpi = dpi)

}
