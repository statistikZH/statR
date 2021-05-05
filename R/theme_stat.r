#' theme_stat()
#'
#' Stat ZH Theme for ggplot2
#' @inheritParams ggplot2::theme_bw
#' @param axis.label.pos position of x and y axis labels, can be "top", "center", or "bottom". Defaults to "top".
#' @keywords theme_stat
#' @export
#' @importFrom ggplot2 theme_minimal theme element_blank element_line unit continuous_scale
#' @examples
#' #'\donttest{
#' \dontrun{
#' ggplot(mpg, aes(class)) +
#' geom_bar() +
#' theme_stat(base_size = 12, axis.label.pos = "center") +
#' labs(title = "Title")

# Function
theme_stat <- function(base_size = 11, axis.label.pos = "top"){

  palette <- RColorBrewer::brewer.pal("Greys", n=9)
  color.grid = palette[5]
  color.title = palette[9]
  color.axis = palette[7]

  ## Achsenposition
  if(axis.label.pos == "top") {
    vjust.x <- 0
    hjust.x <- 1
    vjust.y <- 1
    hjust.y <- 1
  } else if(axis.label.pos == "center") {
    vjust.x <- 0
    hjust.x <- 0.5
    vjust.y <- 1
    hjust.y <- 0.5
  } else if(axis.label.pos == "bottom"){
    vjust.x <- 0
    hjust.x <- 0
    vjust.y <- 1
    hjust.y <- 0
  }

  ggplot2::theme_minimal(base_family = "arial") +
    # TEXTE
    ggplot2::theme(
      text = ggplot2::element_text(size = base_size, color = color.axis, face = "plain", family = "arial"),

      plot.title = ggplot2::element_text(size = base_size*4/3, colour = color.title, face = "bold", family="arialblack"),
      plot.subtitle = ggplot2::element_text(size = base_size, color = color.title, face = "plain", family = "arial"),
      plot.caption = ggplot2::element_text(size = base_size, color = color.title, face = "plain", family = "arial", hjust = 0),

      strip.text = ggplot2::element_text(size = base_size, color = color.axis, face = "plain", family = "arial"),

      axis.title = ggplot2::element_text(size = base_size, color = color.axis, face = "plain", family = "arial"),
      axis.text = ggplot2::element_text(size = base_size, color = color.axis, face = "plain", family = "arial"),

      legend.title = ggplot2::element_text(size = base_size, color = color.axis, face = "plain", family = "arial"),
      legend.text = ggplot2::element_text(size = base_size, color = color.axis, face = "plain", family = "arial")
    ) +

    # ACHSEN

    ggplot2::theme(
      axis.title.x = ggplot2::element_text(color=color.axis, vjust= vjust.x, hjust = hjust.x),
      axis.title.y = ggplot2::element_text(color=color.axis, vjust= vjust.y, hjust = hjust.y),
      axis.line = ggplot2::element_line(colour = color.axis, size = 0.2),
      axis.line.x = ggplot2::element_line(colour = color.axis, size = 0.25),
      axis.line.y = ggplot2::element_blank(),
      axis.ticks = ggplot2::element_line(colour = color.axis, size = 0.25),
      axis.ticks.x = ggplot2::element_line(colour = color.axis, size = 0.25),
      axis.ticks.y = ggplot2::element_blank(),
      axis.ticks.length = unit(0.1, "cm")
    )+

    # GITTERNETZLINIEN
    ggplot2::theme(
      panel.grid.minor = ggplot2::element_blank(),
      panel.grid.major = ggplot2::element_line(colour = color.grid,  size = 0.2),
      panel.grid.major.x = ggplot2::element_blank(), panel.grid.minor.x = ggplot2::element_blank()

    ) 	+

    # PANELS
    ggplot2::theme(
      panel.spacing.x = unit(15, "pt"),
      panel.spacing.y = unit(15, "pt"),
      strip.background = ggplot2::element_blank()

    )+
    # LEGEND
    ggplot2::theme(
      legend.box.spacing = unit(0, "cm"),
      legend.box.margin = ggplot2::margin(0, 0, 1, -1, "mm"),
      legend.position = "top",
      legend.justification = "left",
      # legend.spacing = unit(c(0, 0, 0, 0), "mm"),
      legend.key.width = unit(4, "mm"),
      legend.key.height = unit(3, "mm")
    )+

    # PLOT MARGINS
    ggplot2::theme(plot.margin = unit(c(0.3, 0.3, 0.3, 0.3), "cm"))
}

