#' Style definitions
#'
#' Functions which generate style objects. Not intended to be called directly by the user.
#' @keywords internal
#' @noRd

mainTitleStyle <- function(){
  openxlsx::createStyle(fontSize = 20, textDecoration = "bold", fontName = "Arial",
                        halign = "left")
}

style_title <- function(){
  openxlsx::createStyle(fontSize = 14, textDecoration = "bold", fontName = "Arial")
}

subtitleStyle <- function(){
  openxlsx::createStyle(fontSize = 11, textDecoration = "bold", fontName = "Arial",
                        halign = "left")
}

style_subtitle <- function(){
  openxlsx::createStyle(fontSize = 11, fontName = "Calibri")
}

style_subtitle2 <- function(){
  openxlsx::createStyle(fontSize = 12, textDecoration = "bold", fontName = "Arial")
}

style_subtitle3 <- function(){
  openxlsx::createStyle(fontSize = 12, textDecoration = "italic", fontName = "Arial")
}

hyperlinkStyle <- function(){
  openxlsx::createStyle(fontSize = 11, fontName = "Calibri", fontColour = "blue",
                        textDecoration = "underline")
}

style_header <- function(){
  openxlsx::createStyle(fontSize = 12, fontName = "Calibri", fontColour = "#000000",
    halign = "left", border="Bottom", borderColour = "#009ee0", textDecoration = "bold")
}


# Linien --------------
style_headerline <- function(){
  openxlsx::createStyle(border = "Bottom", borderColour = "#009ee0",
                        borderStyle = getOption("openxlsx.borderStyle", "thick"))
}

style_bottomline <- function(){
  openxlsx::createStyle(border="Bottom", borderColour = "#009ee0")
}

style_leftline <- function(){
  openxlsx::createStyle(border = "Left", borderColour = "#009ee0")
}

# Linewrap text ---------
style_wrap <- function() {
  openxlsx::createStyle(wrapText = TRUE)
}
