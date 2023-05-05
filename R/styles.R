#' Style definitions
#'
#' Description
#'
#' @keywords styles
#' @examples
#'\donttest{
#' \dontrun{
#'}
#'}

mainTitleStyle <- function(){
  openxlsx::createStyle(fontSize = 20,
                        textDecoration = "bold",
                        fontName = "Arial",
                        halign = "left")
}

subtitleStyle <- function(){
  openxlsx::createStyle(fontSize=11,
                        textDecoration="bold",
                        fontName="Arial",
                        halign = "left")
}

hyperlinkStyle <- function(){
  openxlsx::createStyle(fontName = "Calibri",
                        fontSize = 11,
                        fontColour = "blue",
                        textDecoration = "underline")
}
