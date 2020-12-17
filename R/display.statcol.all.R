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

  gg1<-ggplot(df, aes(x=seq_along(col), y=0, color=col))+
    geom_point(size=20,shape=15)+
    facet_wrap(~palette, ncol=3,scales="free")+
    guides(color=FALSE)+
    scale_color_manual(values = df$col)+
    theme_minimal()+
    theme(axis.ticks=element_blank(),axis.text=element_blank(),axis.title=element_blank(), panel.grid  =element_blank())

  gg1

}
