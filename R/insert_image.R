# function insert_worksheet_image
#' Title
#'
#' @param image
#' @param workbook
#' @param sheetname
#' @param startrow
#' @param startcol
#' @param width
#' @param height
#'
#' @return
#' @export
#'
#' @examples


insert_worksheet_image = function(image
                                  ,workbook
                                  ,sheetname
                                  ,startrow
                                  ,startcol
                                  ,width
                                  ,height
){
  openxlsx::addWorksheet(workbook
                         ,sheetname
                         ,gridLines = FALSE
  )
  print(image)
  openxlsx::insertPlot(workbook
                       ,sheetname
                       ,width = width
                       ,height = height
                       ,startRow = startrow
                       ,startCol = startcol
                       ,fileType = "png"
                       ,units = "in"
  )
}

