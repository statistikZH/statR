#' flush_left()
#'
#' @description Function to flush title, subtitle and caption to the
#'   lefthand side of the ggplot2 graphics device
#' @param g ggplot object
#' @examples
#' plt <- ggplot2::ggplot(mtcars, ggplot2::aes(x = cyl)) +
#'   ggplot2::geom_bar() +
#'   ggplot2::labs(title = "Cylinders", subtitle = "Bar chart",
#'                 caption = "mtcars")
#'
#' flush_left(plt)
#' @keywords flush_left
#' @export
flush_left <- function(g){

  xout <- ggplot2::ggplotGrob(g)
  xout$layout$l[xout$layout$name == "title"] <- 1
  xout$layout$l[xout$layout$name == "subtitle"] <- 1
  xout$layout$l[xout$layout$name == "caption"] <- 1
  gridExtra::grid.arrange(xout)
}

