#' flush_left()
#'
#' Function to flush title, subtitle and caption to the lefthand side of the
#' graphics device
#' @param g ggplot object
#' @keywords flush_left, ggplot
#' @importFrom ggplot2 ggplotGrob labs aes
#' @importFrom gridExtra grid.arrange
#' @export
#' @examples
#' plt <- ggplot2::ggplot(mtcars, ggplot2::aes(x = cyl)) +
#'   ggplot2::geom_bar() +
#'   ggplot2::labs(title = "Cylinders", subtitle = "Bar chart",
#'     caption = "mtcars")
#'
#' flush_left(plt)
flush_left <- function(g){

  xout <- ggplot2::ggplotGrob(g)
  xout$layout$l[xout$layout$name == "title"] <- 1
  xout$layout$l[xout$layout$name == "subtitle"] <- 1
  xout$layout$l[xout$layout$name == "caption"] <- 1
  gridExtra::grid.arrange(xout)
}

