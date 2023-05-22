#' insert_worksheet_image()
#'
#' @description Insert an image into a worksheet
#' @note
#' The function does not write the result into a .xlsx file. A separate call
#' to openxlsx::saveWorkbook() is required. A temporary file is created when
#' image refers to a gg, ggplot or histogram object. This file is deleted
#' afterwards.
#' @param image image or plot
#' @param wb workbook object to write new worksheet in
#' @param sheetname  name of the sheet tab
#' @param startrow row coordinate of upper left corner of figure
#' @param startcol column coordinate of upper left corner of figure
#' @param width width of figure
#' @param height height of figure
#' @export
#' @importFrom grDevices png dev.off
#' @importFrom methods is
#' @importFrom openxlsx addWorksheet insertImage
#' @importFrom ggplot2 ggsave
#' @keywords insert_worksheet_image
#' @examples
#' library(ggplot2)
#' wb <- openxlsx::createWorkbook()
#'
#' # Add a geom_histogram()
#' #--------------
#' figure1 <- ggplot(mtcars, aes(x = disp)) +
#'   geom_histogram()
#'
#' insert_worksheet_image(image = figure1, wb, sheetname = "ggplot image",
#'   width = 3.5, height = 5.5)
#'
#' # Add a base histogram
#' #--------------
#' base_hist <- hist(mtcars$disp)
#'
#' insert_worksheet_image(image = base_hist, wb, sheetname = "Base histogram",
#'   width = 3.5, height = 5.5)
#'
#' # Insert an existing image file
#' #--------------
#' image_path <- paste0(path.package("statR"), "/extdata/Stempel_Kanton_ZH.png")
#'
#' insert_worksheet_image(image = image_path, wb, sheetname = "Existing image",
#'   width = 3.5, height = 5.5)
#'
#' # Export workbook
#' #--------------
#' \dontrun{
#' openxlsx::saveWorkbook(wb,"insert_worksheet_image.xlsx")
#' }
#'
insert_worksheet_image <- function(image, wb, sheetname, startrow = 3,
                                   startcol = 3, width, height){

  if (is(image, "character")){
    image_path <- image

  } else if (checkImplementedPlotType(image)){

    # Allocate temporary file of type png to save the image to
    image_path <- tempfile(fileext = ".png")

    # Handle case input object is of class histogram
    if (is(image, "histogram")){
      grDevices::png(image_path, width = width, height = height, units = "in",
                     res = 300)
      plot(image)
      grDevices::dev.off()

    } else {

      # Handle case input object is of class gg or ggplot
      ggplot2::ggsave(image_path, plot = image, width = width, height = height,
                      dpi = 300, device = "png")
    }

  } else {
    stop("Plot muss als ggplot Objekt oder als Filepath vorliegen.")
  }

  # Add new worksheet with image ----------
  if (file.exists(image_path)){
    openxlsx::addWorksheet(wb, sheetName = sheetname, gridLines = FALSE)

    openxlsx::insertImage(wb, sheet = sheetname, file = image_path,
                          width = width, height = height, startRow = startrow,
                          startCol = startcol, units = "in", dpi = 300)

    if (!is.character(image)){
      unlink(image_path)
    }
  } else {
    warning("Image not found.")
  }
}
