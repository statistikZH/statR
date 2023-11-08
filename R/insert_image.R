#' insert_worksheet_image()
#'
#' @description Inserts images into Workbooks. Can either be used to add images
#'   into existing worksheets, or to create new ones with a single image.
#'   Can be configured to add title, source, and metadata.
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
insert_worksheet_image <- function(
    wb, sheetname, image, title = NULL, source = NULL, metadata = NULL,
    width = NULL, height = NULL, startrow = 3, startcol = 3, units = "in", dpi = 300) {

  # If image is.null, pass
  if (is.null(image)) return()

  else if (checkImplementedPlotType(image)) {
    image_path <- ifelse(is.character(image), image, tempfile(fileext = ".png"))

  } else {
    stop("Plot muss als ggplot Objekt oder als Filepath vorliegen.")
  }


  # Try to fill in values if not provided
  if (is.null(title))
    title <- extract_attribute(image, "title")

  if (is.null(source))
    source <- extract_attribute(image, "source")

  if (is.null(metadata))
    metadata <- extract_attribute(image, "metadata")

  if (missing(height) || is.null(height))
    height <- extract_attribute(image, "plot_height", TRUE)

  if (missing(width) || is.null(width))
    width <- extract_attribute(image, "plot_width", TRUE)


  # Create png of plot ----------
  # - histogram (a legacy option.)
  if (is(image, "histogram")) {
    grDevices::png(image_path, width = width, height = height, units = units,
                   res = dpi)
    plot(image)
    grDevices::dev.off()
  }

  else if (inherits(image, c("gg", "ggplot"))) {
    ggplot2::ggsave(image_path, plot = image, device = "png",
                    width = width, height = height, units = units, dpi = dpi)
  }


  # If file not found at image_path, warn and pass
  if (!file.exists(image_path)) {
    warning("Image not found.")
    return()
  }

  # Add worksheet if not already defined ---------
  if (!(sheetname %in% names(wb))) {
    openxlsx::addWorksheet(wb, sheetName = sheetname, gridLines = FALSE)
  }

  header_rows <- c()
  if (!is.null(title) && !all(is.na(title))) {
    openxlsx::writeData(
      wb, sheetname, title, startCol = startcol,
      startRow = startrow, name = paste(sheetname, "imgtitle", sep = "_"))
    openxlsx::addStyle(wb, sheetname, style_title(), startrow, startcol)

    header_rows <- c(header_rows, namedRegionRowExtent(wb, sheetname, "imgtitle"))
    startrow <- max(header_rows) + 1
  }

  if (!is.null(source) && !all(is.na(source))) {
    openxlsx::writeData(
      wb, sheetname, inputHelperSource(source),
      startRow = startrow, startCol = startcol,
      name = paste(sheetname, "imgsource", sep = "_"))

    openxlsx::addStyle(wb, sheetname, style_subtitle(), startrow, startcol,
                       stack = TRUE, gridExpand = TRUE)
    header_rows <- c(header_rows, namedRegionRowExtent(wb, sheetname, "imgsource"))
    startrow <- max(header_rows) + 1
  }

  if (!is.null(metadata) && !all(is.na(metadata))) {
    openxlsx::writeData(
      wb, sheetname, inputHelperMetadata(metadata),
      startRow = startrow, startCol = startcol,
      name = paste(sheetname, "imgmetadata", sep = "_"))

    openxlsx::addStyle(
      wb, sheetname, style_subtitle(), startrow, startcol,
      stack = TRUE, gridExpand = TRUE)

    header_rows <- c(header_rows, namedRegionRowExtent(wb, sheetname, "imgmetadata"))
    startrow <- max(header_rows) + 1
  }

  if (length(header_rows) > 0) {
    purrr::walk(header_rows,
              ~openxlsx::mergeCells(wb, sheetname, cols = startcol + 0:17, rows = .))
    openxlsx::addStyle(wb, sheetname, style_wrap(),
                       header_rows, startcol,
                       stack = TRUE, gridExpand = TRUE)
  }

  # Insert image ---------
  openxlsx::insertImage(wb, sheet = sheetname, file = image_path,
                        width = width, height = height, startRow = startrow,
                        startCol = startcol, units = units, dpi = dpi)

}


