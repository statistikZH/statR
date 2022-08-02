#' theme_stat()
#'
#' This ggplot2 theme is based on ggplot2::theme_minimal(). It controls the non-data related characteristics of a plot (e.g., the font type).
#' On top of that, the font size, the major and minor grid lines, axis lines, axis ticks and the axis label positions can be specified.
#'
#' To use this theme in a R Markdown generated PDF document, insert `dev="cairo_pdf"` into `knitr::opts_chunk$set()`.
#' @inheritParams ggplot2::theme_minimal
#' @param axis.label.pos position of x and y-axis labels, can be set to "top", "center", or "bottom".
#' @param axis.lines presence of axis lines, can be set to "x", "y", "both", or "none".
#' @param ticks presence of axis ticks, can be set to "x", "y", "both", or "none".
#' @param major.grid.lines presence of major grid lines, can be set to "x", "y", "both", or "none".
#' @param minor.grid.lines presence of minor grid lines, can be set to "x", "y", "both", or "none".
#' @param map whether the theme should be optimized for maps, can be set to TRUE or FALSE.
#' @keywords theme_stat
#' @export
#' @importFrom ggplot2 theme_minimal theme element_blank element_line unit continuous_scale
#' @examples
#'\donttest{
#' \dontrun{
#' library(ggplot2)
#' library(statR)
#'
#' ggplot(mpg, aes(class)) +
#' geom_bar() +
#' theme_stat() +
#' labs(title = "Title")
#' }}

theme_stat <- function(base_size = 11, axis.label.pos = "top", axis.lines = "x",
                       ticks = "x", major.grid.lines = "y", minor.grid.lines = "none",
                       map = FALSE){

  palette <- RColorBrewer::brewer.pal("Greys", n=9)
  color.grid = palette[5]
  color.title = palette[9]
  color.axis = palette[7]

  theme_var <- ggplot2::theme_minimal(base_family = "Arial") +
    # TEXTE
    ggplot2::theme(
      text = ggplot2::element_text(size = base_size, color = color.axis, face = "plain", family = "Arial"),

      plot.title = ggplot2::element_text(size = base_size*4/3, colour = color.title, face = "bold", family="Arial Black"),
      plot.subtitle = ggplot2::element_text(size = base_size, color = color.title, face = "plain", family = "Arial"),
      plot.caption = ggplot2::element_text(size = base_size, color = color.title, face = "plain", family = "Arial", hjust = 0),

      strip.text = ggplot2::element_text(size = base_size, color = color.axis, face = "plain", family = "Arial"),

      axis.title = ggplot2::element_text(size = base_size, color = color.axis, face = "plain", family = "Arial"),
      axis.text = ggplot2::element_text(size = base_size, color = color.axis, face = "plain", family = "Arial"),

      legend.title = ggplot2::element_text(size = base_size, color = color.axis, face = "plain", family = "Arial"),
      legend.text = ggplot2::element_text(size = base_size, color = color.axis, face = "plain", family = "Arial")
    )

  # ACHSEN

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

  theme_var <- theme_var +
    ggplot2::theme(axis.title.x = ggplot2::element_text(color=color.axis, vjust= vjust.x, hjust = hjust.x),
                   axis.title.y = ggplot2::element_text(color=color.axis, vjust= vjust.y, hjust = hjust.y))

  ## Achsenlinien

  if(axis.lines != "none"){

    if(axis.lines == "both") {
      theme_var <- theme_var +
        ggplot2::theme(
          axis.line = ggplot2::element_line(colour = color.axis, size = 0.2),
          axis.line.x = ggplot2::element_line(colour = color.axis, size = 0.25),
          axis.line.y = ggplot2::element_line(colour = color.axis, size = 0.25)
        )
    }

    if (axis.lines == "x") {
      theme_var <- theme_var +
        ggplot2::theme(
          axis.line = ggplot2::element_line(colour = color.axis, size = 0.2),
          axis.line.x = ggplot2::element_line(colour = color.axis, size = 0.25),
          axis.line.y = ggplot2::element_blank()
        )
    }

    if (axis.lines == "y") {
      theme_var <- theme_var +
        ggplot2::theme(
          axis.line = ggplot2::element_line(colour = color.axis, size = 0.2),
          axis.line.y = ggplot2::element_line(colour = color.axis, size = 0.25),
          axis.line.x = ggplot2::element_blank(),
        )
    }

  }

  ## Achsenticks

  if(ticks != "none"){
    if(ticks == "x") {
      theme_var <- theme_var +
        ggplot2::theme(
          #axis.ticks = ggplot2::element_line(colour = color.axis, size = 0.25),
          axis.ticks.x = ggplot2::element_line(colour = color.axis, size = 0.25),
          axis.ticks.length = unit(0.1, "cm")
        )
    }

    if(ticks == "y") {
      theme_var <- theme_var +
        ggplot2::theme(
          #axis.ticks = ggplot2::element_line(colour = color.axis, size = 0.25),
          axis.ticks.y = ggplot2::element_line(colour = color.axis, size = 0.25),
          axis.ticks.length = unit(0.1, "cm")
        )
    }

    if(ticks == "both") {
      theme_var <- theme_var +
        ggplot2::theme(
          #axis.ticks = ggplot2::element_line(colour = color.axis, size = 0.25),
          axis.ticks.x = ggplot2::element_line(colour = color.axis, size = 0.25),
          axis.ticks.y = ggplot2::element_line(colour = color.axis, size = 0.25),
          axis.ticks.length = unit(0.1, "cm")
        )
    }
  }

  # GITTERNETZLINIEN

  # keine Gitternetzlinien
  theme_var <- theme_var +
    theme(panel.grid.major.x = ggplot2::element_blank(),
          panel.grid.minor.x = ggplot2::element_blank(),
          panel.grid.major.y = ggplot2::element_blank(),
          panel.grid.minor.y = ggplot2::element_blank())

  # kleine Gitternetzlinien
  if(minor.grid.lines != "none"){

    if(minor.grid.lines == "x") {
      theme_var <- theme_var +
        ggplot2::theme(
          panel.grid.minor.x = ggplot2::element_line(colour = color.grid,  size = 0.1)
        )
    }

    if(minor.grid.lines == "y") {
      theme_var <- theme_var +
        ggplot2::theme(
          panel.grid.minor.y = ggplot2::element_line(colour = color.grid,  size = 0.1)
        )
    }

    if(minor.grid.lines == "both") {
      theme_var <- theme_var +
        ggplot2::theme(
          panel.grid.minor.y = ggplot2::element_line(colour = color.grid,  size = 0.1),
          panel.grid.minor.x = ggplot2::element_line(colour = color.grid,  size = 0.1)
        )
    }
  }

  # grosse Gitternetzlinien
  if(major.grid.lines != "none"){

    if(major.grid.lines == "x") {
      theme_var <- theme_var +
        ggplot2::theme(
          panel.grid.major.x = ggplot2::element_line(colour = color.grid,  size = 0.2)
        )
    }

    if(major.grid.lines == "y") {
      theme_var <- theme_var +
        ggplot2::theme(
          panel.grid.major.y = ggplot2::element_line(colour = color.grid,  size = 0.2)
        )
    }

    if(major.grid.lines == "both") {
      theme_var <- theme_var +
        ggplot2::theme(
          panel.grid.major.y = ggplot2::element_line(colour = color.grid,  size = 0.2),
          panel.grid.major.x = ggplot2::element_line(colour = color.grid,  size = 0.2)
        )
    }
  }

  # PANELS
  theme_var <- theme_var +
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

  # MAP: no grid lines etc. for maps
  if(isTRUE(map)){
    theme_var <- theme_var +
      theme(line = ggplot2::element_blank(),
            axis.line.x = ggplot2::element_blank(),
            axis.ticks.x = ggplot2::element_blank(),
            axis.line.y = ggplot2::element_blank(),
            axis.ticks.y = ggplot2::element_blank(),
            axis.text = ggplot2::element_blank(),
            axis.title = ggplot2::element_blank(),
            panel.grid.major = ggplot2::element_blank(),
            panel.grid.minor = ggplot2::element_blank()) +
      theme(legend.position ='right',
            legend.title = ggplot2::element_text(size = base_size),
            legend.text = ggplot2::element_text(size = base_size))
  }
  theme_var
}

