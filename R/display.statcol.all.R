# display.statcol.all()

#' display.statcol.all()
#'
#' Preview all Stat ZH palettes stored in the 'zhpal'-object
#' @keywords display.statcol.all
#' @export
#' @examples
#' display.statcol.all()
#'

# Function

display.statcol.all<-function() {

  df <- data.frame(unlist(x$stattheme_data))

  df$pal <- row.names(df)

  colnames(df) <- c("col","palette")

  df$palette<-gsub('(.mid|.low|.high|.high|\\d$)','', df$palette)
  df$palette<-gsub('\\d$','', df$palette)

  gg1<-ggplot2::ggplot(df, ggplot2::aes(x=seq_along(col), y=0, color=col))+
    ggplot2::geom_point(size=20,shape=15)+
    ggplot2::facet_wrap(~palette, ncol=3,scales="free")+
    ggplot2::guides(color=FALSE)+
    ggplot2::scale_color_manual(values = df$col)+
    ggplot2::theme_minimal()+
    ggplot2::theme(axis.ticks=ggplot2::element_blank(),
                   axis.text=ggplot2::element_blank(),
                   axis.title=ggplot2::element_blank(),
                   panel.grid=ggplot2::element_blank())

  gg1

}
