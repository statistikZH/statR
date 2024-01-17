#' Insert image into workbooks
#'
#' @description Inserts an image into a new worksheet, optionally with title,
#'   source, and metadata. Images must be provided either as a ggplot object
#'   or as a path to the image file.
#' @note The function does not write the result into a .xlsx file. A separate
#'   call to openxlsx::saveWorkbook() is required. A temporary file is created
#'   for inputs of type gg, ggplot or histogram object at path given by
#'   `tempfile()`.
#' @param wb workbook object to write new worksheet in
#' @param sheetname Name of the sheet where the image should be inserted
#' @param image Image, either a ggplot object or the path to an existing image.
#' @param title Title of the image. Can be NULL
#' @param source Source associated with the image. Can be NULL.
#' @param metadata Metadata associated with the image. Can be NULL.
#' @param startrow row coordinate of upper left corner of figure
#' @param startcol column coordinate of upper left corner of figure
#' @param width width of figure
#' @param height height of figure
#' @param units unit of measurement (default: in)
#' @param dpi image resolution (default: 300)
#'
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
insert_worksheet_image <- function(wb, sheetname, image, title = NULL,
                                   source = NULL, metadata = NULL, width = NULL,
                                   height = NULL, startrow = 3, startcol = 3,
                                   units = "in", dpi = 300) {

  for (value in c("title", "source", "metadata")) {
    if (is.null(eval(as.name(value)))) {
      assign(value, extract_attribute(image, value))
    }
  }

  if (is.null(width)) width <- extract_attribute(image, "plot_width", TRUE)
  if (is.null(height)) height <- extract_attribute(image, "plot_height", TRUE)

  # If image is.null, pass
  if (is.null(image)) return()

  if (checkImplementedPlotType(image)) {
    image_path <- ifelse(is.character(image), image, tempfile(fileext = ".png"))

  } else {
    stop("Plot muss ein ggplot Objekt oder Dateipfad sein.")
  }

  if (inherits(image, c("gg", "ggplot"))) {
    ggplot2::ggsave(image_path, plot = image, device = "png",
                    width = width, height = height, units = units, dpi = dpi)
  }

  # If file not found at image_path, warn and pass
  if (!file.exists(image_path)) {
    warning("Datei ", image_path, " nicht gefunden.")
    return()
  }

  # Add worksheet if not already defined ---------
  if (!(sheetname %in% names(wb))) {
    openxlsx::addWorksheet(wb, sheetName = sheetname, gridLines = FALSE)
  }

  if (is.character(title)) {
    writeText(wb, sheetname, title, startrow, startcol + 0:17, style_title(), "imgtitle")
    startrow <- namedRegionLastRow(wb, sheetname, "imgtitle") + 1
  }

  if (is.character(source)) {
    writeText(wb, sheetname, source, startrow, startcol + 0:17, style_subtitle(), "imgsource")
    startrow <- namedRegionLastRow(wb, sheetname, "imgsource") + 1
  }

  if (is.character(metadata)) {
    writeText(wb, sheetname, metadata, startrow, startcol + 0:17, style_subtitle(), "imgmetadata")
    startrow <- namedRegionLastRow(wb, sheetname, "imgmetadata") + 1
  }

  # Insert image ---------
  openxlsx::insertImage(wb, sheet = sheetname, file = image_path,
                        width = width, height = height, startRow = startrow,
                        startCol = startcol, units = units, dpi = dpi)
}
