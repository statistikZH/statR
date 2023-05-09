#' insert_worksheet_image()
#'
#' insert an image into a single worksheet
#' @description
#' The import
#' @note
#' The function does not write the result into a .xlsx file. A separate call
#' to openxlsx::saveWorkbook() is required.
#' @param image image or plot
#' @param workbook workbook object to write new worksheet in
#' @param sheetname  name of the sheet tab
#' @param startrow row coordinate of upper left corner of figure
#' @param startcol column coordinate of upper left corner of figure
#' @param width width of figure
#' @param height height of figure
#' @importFrom grDevices png dev.off
#' @keywords internal
#' @examples
#' # ggplot object
#' #--------------
#' wb <- openxlsx::createWorkbook()
#'
#' # Add a geom_histogram()
#'
#' figure1 <- ggplot(mtcars, aes(x = disp)) + geom_histogram()
#'
#' insert_worksheet_image(image = figure1, wb, sheetname = "ggplot image",
#'   width = 3.5, height = 5.5)
#'
#' # Add a base histogram
#' base_hist <- histogram(mtcars$disp)
#'
#' insert_worksheet_image(image = base_hist, wb, sheetname = "Base histogram",
#'   width = 3.5, height = 5.5)
#'
#' # Insert an existing image file
#' image_path <- paste0(path.package("statR"), "/extdata/Stempel_Kanton_ZH.png")
#'
#' insert_worksheet_image(image = image_path, wb, sheetname = "Existing image",
#'   width = 3.5, height = 5.5)
#'
#' \dontrun{
#' openxlsx::saveWorkbook(export,"insert_worksheet_image.xlsx")
#' }
#'
insert_worksheet_image = function(
  image,
  wb,
  sheetname,
  width,
  height
){
  openxlsx::addWorksheet(
    wb,
    sheetName = sheetname,
    gridLines = FALSE)

  if (class(image) == "character" ){
    image_path <- image

  } else if (class(image) %in% c("gg", "ggplot", "histogram")){

    # Allocate temporary file of type png to save the image to
    image_path <- tempfile(fileext = ".png")

    # Handle case input object is of class histogram
    if (class(image) == "histogram"){
      png(
        image_path,
        width = width,
        height = height,
        units = "in")
      plot(image)
      dev.off()

    } else {

      # Handle case input object is of class gg or ggplot
      ggplot2::ggsave(
        image_path,
        plot = image,
        width = width,
        height = height,
        dpi = 300,
        device = "png")
    }

  } else {
    stop("Plot muss als ggplot Objekt oder als Filepath vorliegen.")
  }

  # Insert image
  openxlsx::insertImage(
    wb = wb,
    sheet = sheetname,
    file = image_path,
    width = width,
    height = height,
    startRow = 3,
    startCol = 3,
    units = "in",
    dpi = 300
  )

  # Delete temporary file if necessary
  if (!is.character(image)){
    unlink(image_path)
  }
}
