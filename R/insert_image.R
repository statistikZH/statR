#' insert an image into a single worksheet
#' insert_image
#'
#' @param image image or plot
#' @param workbook workbook object to write new worksheet in
#' @param sheetname  name of the sheet tab
#' @param startrow row coordinate of upper left corner of figure
#' @param startcol column coordinate of upper left corner of figure
#' @param width width of figure
#' @param height height of figure
#' @noRd
#' @examples
#' # example
#' \dontrun{
#' export <- openxlsx::createWorkbook("export")
#'
#' insert_worksheet_image(image=plot(x = mtcars$wt, y = mtcars$mpg),
#'                        export,
#'                        "image",
#'                        startrow=2,
#'                        startcol=2,
#'                        width=3.5,
#'                        height=5.5
#' )
#'openxlsx::saveWorkbook(export,"insert_worksheet_image.xlsx")
#'
#'}

insert_worksheet_image = function(image,
                                  wb,
                                  sheetname,
                                  width,
                                  height){

  openxlsx::addWorksheet(wb,
                         sheetname,
                         gridLines = FALSE
  )

  temp <- tempfile(fileext = ".png")
  #p2 <- image
  ggplot2::ggsave(
    temp,
    plot = image,
    width = width,
    height = height,
    dpi = 300,
    device = "png"
    )

  openxlsx::insertImage(
    wb = wb,
    sheet = sheetname,
    file = temp,
    width = width,
    height = height,
    startRow = 3,
    startCol = 3,
    units = "in",
    dpi = 300
  )

  file.remove(temp)
}

# # example
# export <- openxlsx::createWorkbook("export")
#
# insert_worksheet_image(image=plot(x = mtcars$wt, y = mtcars$mpg)
# ,export
# ,"image"
# ,startrow=2
# ,startcol=2
# ,width=3.5
# ,height=5.5
# )
#
# openxlsx::saveWorkbook(export,"insert_worksheet_image.xlsx")
