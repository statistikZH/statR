# ------- Farbpaletten basierend auf Kantons-CI Farben ----------------

library(scales)

#################### STAT ZH Farbpaletten

#' Data used by the stat zh palettes
#'
#' @format A \code{list}.
#' @export
#'
stattheme_data <- {

  x <- list()

  #standard
  x$stattheme_data$zhcd<- c("#009ee0", "#0076bd", "#885ea0", "#e30059", "#e2001a","#ffcc00", "#00a1a3","#3ea743")

  ### 1. Ordinale / kategoriale Paletten

  #Farbmatrixpaletten (horizontal)

  #standard
  x$stattheme_data$zhdarker<- c("#407B9F", "#409292", "#5F9463", "#B2A558", "#B77346", "#B24A4C", "#857091")

  x$stattheme_data$zhdark<- c("#3F8AB4", "#3FA5A5", "#65A96A", "#CEBE5B", "#D27F46", "#CD4B4E", "#9479A4")

  x$stattheme_data$zh<- c("#3F98CC", "#3FB8B9", "#6DBD72", "#EBD760", "#F08D47", "#ea4d51", "#A585B7")

  x$stattheme_data$zhlight<- c("#5FA9D3", "#5FC3C3","#84C788", "#EDDD79", "#F29F64", "#EC686C", "#B399C3")

  x$stattheme_data$zhpastel <- c( "#7FB9DD", "#7FCECE", "#9DD2A0", "#F1E494", "#F4B283", "#F1868B", "#C3ACCE")

  x$stattheme_data$zhextralight <- c("#9FCBE5", "#9FDADA" ,"#B4DEB8", "#F3EBAE", "#F7C5A1", "#F3A5A7", "#D2C1DA")

  x$stattheme_data$zhultralight <- c("#BFDCED" ,"#BFE6E6", "#CDE9D0" ,"#F8F1CA" ,"#F9D8C0", "#F8C3C4", "#E0D6E6")

  #x$stattheme_data$zhpaired16<-c("#4684AA","#8CBFDD", "#469A9A","#8CD0D1", "#669F6A","#A5D3A9", "#CCAA1A","#F8DD6A", "#C37D4C" ,"#F1B892","#BE4654","#EC8C99", "#8D7899","#C6B4CE","#1B252B","#767C7F")

  x$stattheme_data$zhpaired<-c("#3f8ab4","#9FCBE5", "#3fa5a5","#9FDADA", "#65A96A","#B4DEB8", "#CEBE5B","#F3EBAE", "#D27F46" ,"#F7C5A1","#CD4B4E","#F3A5A7", "#9479A4","#D2C1DA","#1B252B","#767C7F")

  #Farbmatrixpaletten (vertikal)

  x$stattheme_data$zhblue <- c("#407B9F","#3F8AB4","#3F98CC","#5FA9D3","#7FB9DD","#9FCBE5","#BFDCED")

  x$stattheme_data$zhbluemh<- c('#386a87','#528489','#699b8f','#7fb19d','#96c3b0','#add3cb','#c5dfed')

  # x$stattheme_data$zhblue <- c('#407b9f','#4c97c1','#71b2d8','#a0cbe5','#d0e5f2')

  #x$stattheme_data$zhblue <- c('#407b9f','#71b2d8','#d0e5f2')

  x$stattheme_data$zhgreen <-c("#527F54", "#669F6A", "#7ABF7F","#90C993","#A5D3A9","#BBDFBF","#D2E9D3")

  x$stattheme_data$zhturqoise <-c("#409292", "#3FA5A5", "#3FB8B8","#5FC3C3","#7FCECE","#9FDADA","#BFE6E6")

  x$stattheme_data$zhyellow<-c("#B2A558", "#CEBE5B", "#EBD760","#EDDD79","#F1E494","#F3EBAE","#F8F1CA")

  x$stattheme_data$zhorange <-c("#B77346", "#D27F46", "#F08D47","#F29f64","#f4b283","#f7c5a1","#f9d8c0")

  x$stattheme_data$zhred <- c('#983844','#b74451','#d84f60','#e76c7b','#ec8c99','#f1a8b2','#f6c5cb')

  x$stattheme_data$zhviolet <-c('#715f7a','#857190','#9a83a7','#af96bc','#c0acc9','#d0c2d8','#e3d9e6')

  x$stattheme_data$zhvioletmh<-c('#715f7a','#877375','#9a8874','#ae9e77','#c1b284','#d4c6a1','#e3d9e6')

  #diagonal palette

  #x$stattheme_data$zhdiagonal <- c("#407B9F", "#3FA5A5" ,"#6CBD72", "#EDDD79", "#F4B283", "#F3A5A7" ,"#E0D6E6")

  x$stattheme_data$zhdiagonal <- c("#407B9F", "#3FA5A5" ,"#6CBD72", "#EDDD79", "#F4B283", "#F3A5A7" ,"#E0D6E6","#857091")

  #zhBlueYellowGrey

  x$stattheme_data$zhbyg<- c( "#003B5E", "#006EB1" ,"#5FA9D5","#FFE16F","#C2D37B", "#86CEB8", "#56BFB9", "#9DB99E", "#656564")

  x$stattheme_data$zhbyglight<- c("#2A5A78", "#2A85BE", "#79B7DC" ,"#FFE586" ,"#CCD991", "#99D6C3" ,"#72C9C4", "#ACC4AD" ,"#7E7E7D")

  #zhlake

  x$stattheme_data$zhlake<-c("#002338","#004671","#0076BD","#3F94BA" ,"#7FB2B7", "#BFD0B4","#DAEBCC", "#FFEFB2")

  x$stattheme_data$zhlakelight<-c("#2A4659", "#2A6487", "#2A8CC7", "#5FA5C5", "#93BEC3", "#C9D7BF" ,"#DFEDD3", "#FFF1BE")

  #likert-paletten

  #zhlikert Benchmarking mit Kantonsrot / Kantonsgr?n - '#e2001a', "#3ea743"

  x$stattheme_data$zhlikert5gr <- c( "#F09B7A", "#E93F53","#6EBD72", "#B3DA8A", "#F8F7A2")

  x$stattheme_data$zhlikert5br <- c("#F09B7A","#E93F53","#3F98CD", "#9BC7B7", "#F8F7A2")

  x$stattheme_data$zhlikert6gr <- c("#6EBD72", "#A5D485" ,"#DCEB98" ,"#f2a975", "#EF8872" ,"#E93F53")

  x$stattheme_data$zhlikert6br <- c("#3F98CD" ,"#7FB9DD", "#9FDADA", "#F7C5A1", "#EF8872", "#E93F53")


  #BENCHMARKING-PALETTEN

  x$stattheme_data$bmlikert <- c("#e2001a","#f27644","#fab872","#ffe16f","#9ec359","#3ea743")

  x$stattheme_data$bmtabelle <- c("#ff9faa","#f9bda5","#d6e6b8","#ffe16f","#b9e5bb")

  x$stattheme_data$bmmittelwert <- c("#6DB0D9")

  #zh180
  # x$stattheme_data$zh180 <- c("#2A5A78","#2A5C7B","#2A5E7E","#2A6081","#2A6284","#2B6487","#2B668A","#2B688D","#2B6A90","#2B6B93","#2B6D96","#2B6F99","#2B719C","#2B739F","#2B75A3","#2B77A6","#2B79A9","#2B7BAC","#2B7DAF","#2A7FB3","#2A81B6","#2A83B9","#2A85BC","#2D87BE","#3289C0","#368BC1","#3A8DC2","#3F8FC4","#4292C5","#4694C6","#4A96C8","#4D98C9","#519ACA","#549CCC","#589FCD","#5BA1CE","#5EA3D0","#61A5D1","#64A7D2","#67AAD4","#6AACD5","#6DAED6","#70B0D8","#73B3D9","#76B5DA","#7BB7DA","#85B8D5","#8EBAD0","#96BCCB","#9EBDC6","#A5BFC1","#ACC1BC","#B2C2B6","#B9C4B1","#BFC6AC","#C4C8A7","#CAC9A1","#CFCB9C","#D4CD96","#D9CF91","#DED08B","#E3D285","#E8D47F","#EDD679","#F1D873","#F5DA6D","#FADB66","#FEDD5F","#FDDD61","#FBDD63","#F8DD66","#F6DD68","#F4DD6A","#F2DD6D","#F0DC6F","#EEDC71","#EBDC74","#E9DC76","#E7DC78","#E5DB7A","#E2DB7C","#E0DB7F","#DEDB81","#DBDB83","#D9DB85","#D7DA87","#D4DA89","#D2DA8B","#CFDA8D","#CDDA8F","#CBD992","#C9D994","#C7D996","#C5D999","#C3D99B","#C1D89D","#BFD89F","#BDD8A2","#BBD8A4","#B9D8A6","#B7D8A8","#B4D7AB","#B2D7AD","#B0D7AF","#AED7B1","#ABD7B3","#A9D7B6","#A6D6B8","#A4D6BA","#A1D6BC","#9FD6BE","#9CD6C1","#99D5C3","#98D5C3","#96D4C3","#94D4C3","#93D3C3","#91D3C3","#8FD2C3","#8ED1C3","#8CD1C3","#8AD0C3","#88D0C3","#87CFC3","#85CEC3","#83CEC3","#81CDC3","#7FCDC3","#7ECCC3","#7CCCC3","#7ACBC3","#78CAC3","#76CAC3","#74C9C3","#72C9C3","#74C8C3","#77C8C2","#7BC8C1","#7EC8C0","#81C8BF","#84C7BE","#86C7BD","#89C7BC","#8CC7BB","#8FC6BA","#91C6B9","#94C6B8","#96C6B7","#99C6B6","#9BC5B5","#9EC5B4","#A0C5B3","#A2C5B2","#A5C4B1","#A7C4B0","#A9C4AF","#ABC4AE","#ACC2AD","#AABFAA","#A7BCA8","#A5B9A6","#A3B5A4","#A1B2A1","#9FAF9F","#9DAC9D","#9BA99B","#99A699","#97A296","#949F94","#929C92","#909990","#8E968E","#8C938C","#8A9089","#888D87","#868A85","#848783","#828481","#80817F","#7E7E7D")


  # 2. Definiere Liste mit sequentiellen Paletten bzw. low / high Farben

  #Da zhdarker nicht dunkel genug ist - mische diese mit schwarz
  #> interpolate(colors$zhdarker, color="#000000",degree=3,7)
  #[1] "#2A526A" "#2A6060" "#3F6141" "#766D39" "#794C2D" "#763132" "#584A5F"

  x$stattheme_data$zhbygseq <- x$stattheme_data$zhbyg[c(4,7,1)]

  x$stattheme_data$zhseq <- c(x$stattheme_data$zhultralight[3],x$stattheme_data$zh[2], x$stattheme_data$zhdarker[1],"#2A526A")

  x$stattheme_data$zhlakeseq <-c(x$stattheme_data$zhlake[8],x$stattheme_data$zhlake[3],x$stattheme_data$zhlakelight[1])


  # x$stattheme_data$sequential$zhblueseq$low <- "#C5DFED"

  # x$stattheme_data$sequential$zhblueseq$high  <- "#386A87"

  x$stattheme_data$sequential$zhblueseq$low <- "#BFDCED"

  x$stattheme_data$sequential$zhblueseq$high  <- "#19242D"


  x$stattheme_data$sequential$zhredseq$low<- "#F6C5CB"

  x$stattheme_data$sequential$zhredseq$high <-"#983844"

  x$stattheme_data$sequential$zhvioletseq$high<- "#715F7A"

  x$stattheme_data$sequential$zhvioletseq$low  <- "#E3D9E6"

  x$stattheme_data$sequential$zhbgseq$high<- "#0076BD"

  x$stattheme_data$sequential$zhbgseq$low  <- "#B1DBB3"

  x$stattheme_data$sequential$zhgreenseq$high<- "#527F54"

  x$stattheme_data$sequential$zhgreenseq$low  <- "#D2E9D3"


  #x$stattheme_data$sequential$zhseq$low <- "#E9F3F9"

  #x$stattheme_data$sequential$zhseq$high  <- "#005A90"

  #3. Definiere Liste mit divergierenden Paletten

  #zhmh - multihue

  x$stattheme_data$diverging$zhmhblue$low <- "#F3EBAE"

  x$stattheme_data$diverging$zhmhblue$mid <- "#0076BD"

  x$stattheme_data$diverging$zhmhblue$high  <- "#19242D"


  x$stattheme_data$diverging$zhmhgreen$low <- "#F3EBAE"

  x$stattheme_data$diverging$zhmhgreen$mid <- "#00a1a3"

  x$stattheme_data$diverging$zhmhgreen$high  <- "#19242D"



  x$stattheme_data$diverging$zhbluered$low <- "#e2001a"

  x$stattheme_data$diverging$zhbluered$mid <- "#f0f0f0"

  x$stattheme_data$diverging$zhbluered$high  <- "#0076BD"


  x$stattheme_data$diverging$zhvioletgreen$high <-"#3ea743"

  x$stattheme_data$diverging$zhvioletgreen$mid <- "#f0f0f0"

  x$stattheme_data$diverging$zhvioletgreen$low<- "#885ea0"


  x$stattheme_data$diverging$zhspectral$low <- "#EA5251"

  x$stattheme_data$diverging$zhspectral$mid <- "#FFE57F"

  x$stattheme_data$diverging$zhspectral$high  <- "#3F98CC"

}


# stattheme_pal()

#' stattheme_pal()
#'
#' Stat ZH scales
#' @keywords stattheme_pal
#' @noRd
#'

#Funktion für Palettenauswahl (ord./kat.)

stattheme_pal <- function(palette = "default") {
  if (palette %in% names(x$stattheme_data)) {
    scales::manual_pal(unname( x$stattheme_data[[palette]]))
  } else {
    stop(sprintf("palette %s not a valid statZH palette.", palette))
  }
}


#' Data used by the stat zh palettes
#'
#' @format A \code{list}.
#' @export
#'

zhpal <- x$stattheme_data
