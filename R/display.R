# display()

#' display()
#'
#' preview colors in console
#' @param col a vector of hex-code colors to be displayed
#' @param border border color
#' @param ... further arguments that can be passed to the plot()-function
#' @keywords display
#' @importFrom grDevices "colorRampPalette"
#' @importFrom graphics "rect"
#' @export
#' @examples
#' #display can be used to preview all the palettes included in the 'zhpal'-list:
#' display(zhpal$zhdiagonal)
#' #example with a single hex-code
#' display("#000000")

# Function

#Funktion um Paletten anzuzeigen
display <- function(col, border = "light gray", ...){
  n <- length(col)
  plot(0, 0, type="n", xlim = c(0, 1), ylim = c(0, 1),
       axes = FALSE, xlab = "", ylab = "", ...)
  rect(0:(n-1)/n, 0, 1:n/n, 1, col = col, border = border)
}
