# interpolate2()

#'interpolate2()
#'
#' generate new palettes by interpolation
#' @param palette a vector of colors to be interpolated to the same degree with the same color to generate uniform color palettes
#' @param color a single color (hex-code) to be mixed with the colors in the 'palette'. 
#' @param degree degree of interpolation, range: 1 (slight interpolation) - 7 (strong interpolation).
#' @param number number of colours to be generated.
#' @keywords interpolate2
#' @export
#' @examples
#' #interpolate the 'zhblue'-palette with black and generate 11 new hues.
#' interpolate2(zhpal$zhblue,"#000000",3,11)
#' 

# Function


#interpolate function
interpolate2<-function(palette, color = color, degree=degree,number=number){
  newpalette <- c()
  for (i in palette) {
    intcols <-c(i,color)
    pal = colorRampPalette(intcols)
    #  plotColors(pal, 7)
    intcols<- pal(7)
    newpalette[i] <- intcols[degree]
  } 
  pal = colorRampPalette((newpalette),space="Lab")
  show_col(pal(number))
  pal(number)
}

