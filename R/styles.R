#' Style definitions
#'
#' Functions which generate style objects. Not intended to be called directly
#' by the user.
#' @keywords internal
#' @noRd


### Worksheet no header
style_maintitle <- function() {
  openxlsx::createStyle(fontSize = 20, textDecoration = "bold", halign = "left",
                        fontName = "Arial")
}

style_title <- function() {
  openxlsx::createStyle(fontSize = 14, textDecoration = "bold",
                        fontName = "Arial")
}

style_indextitle <- function() {
  openxlsx::createStyle(fontSize = 10, textDecoration = "bold", halign = "left",
                        fontName = "Arial", valign = "top" )
}

style_subtitle <- function() {
  openxlsx::createStyle(fontSize = 10, textDecoration = NULL,
                        fontName = "Arial", halign = "left", valign = "top",
                        wrapText = TRUE)
}

hyperlinkStyle <- function() {
  openxlsx::createStyle(fontSize = 10, fontName = "Arial",
                        fontColour = "blue", textDecoration = "underline")
}

style_header <- function() {
  openxlsx::createStyle(
    fontSize = 12, fontName = "Arial", halign = "left", fontColour = "#000000",
    border = "Bottom", borderColour = "#009ee0", borderStyle = "medium",
    textDecoration = "bold")
}

# Linien --------------
style_headerline <- function() {
  openxlsx::createStyle(
    border = "Bottom", borderColour = "#009ee0", borderStyle = "medium")
}

style_leftline <- function() {
  openxlsx::createStyle(
    border = "Left", borderColour = "#009ee0", borderStyle = "thin")
}

# Linewrap text ---------
style_wrap <- function() {
  openxlsx::createStyle(wrapText = TRUE)
}
