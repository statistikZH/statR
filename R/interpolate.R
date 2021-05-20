
#'interpolate2()
#'
#' Function to generate new color palettes by interpolation.
#' @param palette a vector of colors to be interpolated
#' @param color color (hex-code) to be mixed with the colors in the 'palette'.
#' @param degree degree of interpolation, range: 1 (slight interpolation) - 7 (strong interpolation).
#' @param number number of colors to be generated.
#' @keywords interpolate2
#' @export
#' @examples
#' #interpolate the 'zhblue'-palette with black and generate 7 new ones
#' interpolate2(palette = zhpal$zhblue, color = c("#000000"), degree = 3, number = 7)
#'

# Function

#interpolate function
interpolate2<-function(palette, color = color, degree=degree,number=number){
  newpalette <- c()
  for (i in palette) {
    intcols <-c(i,color)
    pal = grDevices::colorRampPalette(intcols)
    #  plotColors(pal, 7)
    intcols<- pal(7)
    newpalette[i] <- intcols[degree]
  }
  pal = grDevices::colorRampPalette((newpalette),space="Lab")
  scales::show_col(pal(number))
  pal(number)
}

