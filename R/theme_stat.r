

# ---------------------- BASIEREND AUF THEME_MINIMAL -----------------------
library(RColorBrewer)
library(ggplot2)
library(grid)
# windowsFonts(arial=windowsFont("Arial Fett"))


palette <- brewer.pal("Greys", n=9)
color.grid = palette[5]
color.title = palette[9]
color.axis = palette[7]

txt <- element_text(size = 11, colour = color.axis, face = "plain")
bold_txt <- element_text(size = 11, colour = color.axis, face = "bold")

# theme_stat

#' theme_stat()
#'
#' Stat ZH Theme for ggplot2
#' @inheritParams ggplot2::theme_bw
#' @keywords theme_stat
#' @export
#' @examples
#' theme_stat()
#' 

# Function
theme_stat <- function(base_size = 11
													 # , base_family = "arial"
													 ){
	theme_minimal(base_size = base_size
	              # , base_family = base_family
	              ) +
		theme(
			legend.key = element_blank(),
			strip.background = element_blank(),
			
			legend.position = "bottom",
			
			text = txt,
			plot.title = element_text(size = 14, colour = color.title, face = "bold"),
			strip.text = txt,
			
			axis.title = txt,
			axis.text = txt,
			
			legend.title = bold_txt,
			legend.text = txt ) +
		
		# ACHSEN
		theme(
			axis.title.x = element_text(color=color.axis, vjust= 0),
			axis.title.y = element_text(color=color.axis, vjust= 1.25),
			axis.line = element_line(colour = color.axis, size = 0.5),
			axis.line.x = element_line(colour = color.axis, size = 0.5),
			axis.line.y = element_blank(),
			axis.ticks = element_line(colour = color.axis, size = 0.5),
			axis.ticks.x = element_line(colour = color.axis, size = 0.5),
			axis.ticks.y = element_blank(),
			# axis.ticks.x = element_line(colour = color.axis, size = 0.8),
			axis.ticks.length = unit(0.15, "cm")
		) +
		# theme(plot.title = element_text(hjust = -1))+
    
		# GITTERNETZLINIEN
		theme(panel.grid.minor = element_line(colour = color.grid, linetype = "dotted", size = 0.4),
					panel.grid.major = element_line(colour = color.grid, linetype = "dotted",  size = 0.4)
				, panel.grid.major.x = element_blank(), panel.grid.minor.x = element_blank()
		) 	+
		
		# Plot margins
		theme(plot.margin = unit(c(0.3, 0.3, 0.3, 0.3), "cm")	)
}

