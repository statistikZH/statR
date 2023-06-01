#' display.statcol.all()
#'
#' @description Preview all Stat ZH palettes stored in the 'zhpal'-object
#' @examples
#' display.statcol.all()
#' @keywords display.statcol.all
#' @importFrom ggplot2 ggplot aes geom_point facet_wrap guides theme_minimal
#'                     theme element_blank labs
#' @importFrom rlang .data
#' @export
display.statcol.all <- function(){

  df <- data.frame(unlist(x$stattheme_data))
  df[,"pal_col"] <- row.names(df)
  colnames(df) <- c("col", "pal_col")

  df$palette <- gsub('(.mid|.low|.high|.high|\\d$)', '', df$pal_col)

  df$palette <- gsub('\\d$','', df$palette)

  gg1 <- ggplot2::ggplot(df, ggplot2::aes(x = .data$pal_col, y = 0, color = I(col))) +
    ggplot2::geom_point(size = 20, shape = 15) +
    ggplot2::facet_wrap(~palette, ncol = 3, scales = "free") +
    ggplot2::guides(color = "none") +
    ggplot2::theme_minimal() +
    ggplot2::theme(axis.ticks = ggplot2::element_blank(),
                   axis.text = ggplot2::element_blank()) +
    ggplot2::labs(x = "", y = "")

  print(gg1)
}
