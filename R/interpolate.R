#'interpolate2()
#'
#' Function to generate new color palettes by interpolation.
#' @param palette a vector of colors to be interpolated
#' @param color color (hex-code) to be mixed with the colors in the 'palette'.
#' @param degree degree of interpolation, range: 1 (slight interpolation) - 7 (strong interpolation).
#' @param number number of colors to be generated.
#' @keywords interpolate2
#' @export
#' @importFrom grDevices colorRampPalette
#' @examples
#' #interpolate the 'zhblue'-palette with black and generate 7 new ones
#' interpolate2(palette = zhpal$zhblue, color = c("#000000"), degree = 3,
#'   number = 7)
#'
interpolate2 <- function(palette, color, degree, number){

  if (degree < 1 | degree > 7){
    stop("degree out of range. Provide integer between 1-7.")
  }

  newpalette <- c()
  for (i in palette) {
    intcols <- c(i, color)
    pal <- grDevices::colorRampPalette(intcols)
    intcols <- pal(7)
    newpalette[i] <- intcols[degree]
  }

  pal <- grDevices::colorRampPalette(newpalette, space = "Lab")
  return(pal(number))
}

