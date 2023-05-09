#' Style definitions
#'
#' @Description
#' Functions which generate style objects. Not intended to be called directly by the user.
#' @keywords internal

mainTitleStyle <- function(){
  openxlsx::createStyle(fontSize = 20,
                        textDecoration = "bold",
                        fontName = "Arial",
                        halign = "left")
}

style_title <- function(){
  openxlsx::createStyle(
    fontSize = 14,
    textDecoration = "bold",
    fontName = "Arial"
  )
}

subtitleStyle <- function(){
  openxlsx::createStyle(fontSize = 11,
                        textDecoration = "bold",
                        fontName = "Arial",
                        halign = "left")
}

style_subtitle <- function(){
  openxlsx::createStyle(
    fontSize = 11,
    fontName = "Calibri"
  )
}

hyperlinkStyle <- function(){
  openxlsx::createStyle(fontName = "Calibri",
                        fontSize = 11,
                        fontColour = "blue",
                        textDecoration = "underline")
}

style_header <- function(){
  openxlsx::createStyle(
    fontSize = 12,
    fontName = "Calibri",
    fontColour = "#000000",
    halign = "left",
    border="Bottom",
    borderColour = "#009ee0",
    textDecoration = "bold"
  )
}

style_headerline <- function(){
  openxlsx::createStyle(border = "Bottom",
                        borderColour = "#009ee0",
                        borderStyle = getOption("openxlsx.borderStyle", "thick"))
}

style_leftline <- function(){
  openxlsx::createStyle(
    border = "Left",
    borderColour = "#009ee0"
  )
}

style_wrap <- function() {
  openxlsx::createStyle(wrapText = TRUE)
}
