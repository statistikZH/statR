# insert an image into a single worksheet
#' insert_image
#'
#' @param image image or plot
#' @param workbook workbook object to write new worksheet in
#' @param sheetname  name of the sheet tab
#' @param startrow row coordinate of upper left corner of figure
#' @param startcol column coordinate of upper left corner of figure
#' @param width width of figure
#' @param height height of figure
#' @export
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
                                  workbook,
                                  sheetname,
                                  startrow,
                                  startcol,
                                  width,
                                  height){

  openxlsx::addWorksheet(workbook,
                         sheetname,
                         gridLines = FALSE
  )

  # print(image)

  openxlsx::insertPlot(workbook,
                       sheetname,
                       width = width,
                       height = height,
                       startRow = startrow,
                       startCol = startcol,
                       fileType = "png",
                       units = "in"
  )
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
