#' theme_stat()
#'
#' Stat ZH Theme for ggplot2
#' @inheritParams ggplot2::theme_bw
#' @keywords theme_stat
#' @export
#' @examples
#' @importFrom ggplot2 theme_minimal theme element_blank element_line unit continuous_scale

# Function
theme_stat <- function(base_size = 11){


  # braucht es das hier noch?
  palette <- RColorBrewer::brewer.pal("Greys", n=9)
  color.grid = palette[5]
  color.title = palette[9]
  color.axis = palette[7]

  txt <- ggplot2::element_text(size = 11, colour = color.axis, face = "plain")
  bold_txt <- ggplot2::element_text(size = 11, colour = color.axis, face = "bold")
  # ---------

	ggplot2::theme_minimal(base_size = base_size
	              # , base_family = base_family
	              ) +
    ggplot2::theme(
			legend.key =  ggplot2::element_blank(),
			strip.background =  ggplot2::element_blank(),

			legend.position = "bottom",

			text = txt,
			plot.title =  ggplot2::element_text(size = 14, colour = color.title, face = "bold"),
			strip.text = txt,

			axis.title = txt,
			axis.text = txt,

			legend.title = bold_txt,
			legend.text = txt ) +

		# ACHSEN
    ggplot2::theme(
			axis.title.x = ggplot2::element_text(color=color.axis, vjust= 0),
			axis.title.y = ggplot2::element_text(color=color.axis, vjust= 1.25),
			axis.line = ggplot2::element_line(colour = color.axis, size = 0.5),
			axis.line.x = ggplot2::element_line(colour = color.axis, size = 0.5),
			axis.line.y = element_blank(),
			axis.ticks = ggplot2::element_line(colour = color.axis, size = 0.5),
			axis.ticks.x = ggplot2::element_line(colour = color.axis, size = 0.5),
			axis.ticks.y = element_blank(),
			# axis.ticks.x = ggplot2::element_line(colour = color.axis, size = 0.8),
			axis.ticks.length = unit(0.15, "cm")
		) +
		# theme(plot.title = ggplot2::element_text(hjust = -1))+

		# GITTERNETZLINIEN
		theme(panel.grid.minor = ggplot2::element_line(colour = color.grid, linetype = "dotted", size = 0.4),
					panel.grid.major = ggplot2::element_line(colour = color.grid, linetype = "dotted",  size = 0.4)
				, panel.grid.major.x =  ggplot2::element_blank(), panel.grid.minor.x =  ggplot2::element_blank()
		) 	+

		# Plot margins
    ggplot2::theme(plot.margin = unit(c(0.3, 0.3, 0.3, 0.3), "cm")	)
}

