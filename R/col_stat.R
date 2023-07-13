# ------- Farbpaletten basierend auf Kantons-CI Farben ----------------

#################### STAT ZH Farbpaletten
#' stattheme_data
#'
#' Data used by the stat zh palettes
#'
#' @format A \code{list}.
#' @noRd
#'
stattheme_data <- {

  x <- list()

  # standard
  x$stattheme_data$zhcd <- c("#009ee0", "#0076bd", "#885ea0", "#e30059",
                             "#e2001a", "#ffcc00", "#00a1a3", "#3ea743")

  # Akzentfarben ZH Web
  x$stattheme_data$zhwebaccent <- c("#009ee0", "#0070B4", "#7F3DA7", "#D40053",
                                    "#B01657", "#DC7700", "#00797B", "#1A7F1F")


  ### 1. Ordinale / kategoriale Paletten

  # Farbmatrixpaletten (horizontal)

  x$stattheme_data$zhwebdataviz <- c("#0070B4", "#00407C", "#00797B", "#8A8C00",
                                     "#00544C", "#DC7700", "#D93C1A", "#96170F",
                                     "#D40053", "#B01657", "#7A0049", "#9572D5",
                                     "#54268E", "#CCCCCC", "#949494", "#666666",
                                     "#333333")

  x$stattheme_data$zhdarker <- c("#407B9F", "#409292", "#5F9463", "#B2A558",
                                 "#B77346", "#B24A4C", "#857091")

  x$stattheme_data$zhdark <- c("#3F8AB4", "#3FA5A5", "#65A96A", "#CEBE5B",
                               "#D27F46", "#CD4B4E", "#9479A4")

  x$stattheme_data$zh <- c("#3F98CC", "#3FB8B9", "#6DBD72", "#EBD760",
                           "#F08D47", "#ea4d51", "#A585B7")

  x$stattheme_data$zhlight <- c("#5FA9D3", "#5FC3C3", "#84C788", "#EDDD79",
                                "#F29F64", "#EC686C", "#B399C3")

  x$stattheme_data$zhpastel <- c("#7FB9DD", "#7FCECE", "#9DD2A0", "#F1E494",
                                 "#F4B283", "#F1868B", "#C3ACCE")

  x$stattheme_data$zhextralight <- c("#9FCBE5", "#9FDADA", "#B4DEB8", "#F3EBAE",
                                     "#F7C5A1", "#F3A5A7", "#D2C1DA")

  x$stattheme_data$zhultralight <- c("#BFDCED", "#BFE6E6", "#CDE9D0", "#F8F1CA",
                                     "#F9D8C0", "#F8C3C4", "#E0D6E6")


  x$stattheme_data$zhpaired <- c("#3f8ab4", "#9FCBE5", "#3fa5a5", "#9FDADA",
                                 "#65A96A", "#B4DEB8", "#CEBE5B", "#F3EBAE",
                                 "#D27F46", "#F7C5A1", "#CD4B4E", "#F3A5A7",
                                 "#9479A4", "#D2C1DA", "#1B252B", "#767C7F")


  # Farbmatrixpaletten (vertikal)
  #----------------------------
  x$stattheme_data$zhblue <- c("#407B9F", "#3F8AB4", "#3F98CC", "#5FA9D3",
                               "#7FB9DD", "#9FCBE5", "#BFDCED")

  x$stattheme_data$zhgreen <- c("#527F54", "#669F6A", "#7ABF7F", "#90C993",
                                "#A5D3A9", "#BBDFBF", "#D2E9D3")

  x$stattheme_data$zhorange <- c("#B77346", "#D27F46", "#F08D47", "#F29f64",
                                 "#f4b283", "#f7c5a1", "#f9d8c0")

  x$stattheme_data$zhred <- c("#983844", "#b74451", "#d84f60", "#e76c7b",
                              "#ec8c99", "#f1a8b2", "#f6c5cb")


  # Diagonal palette
  #-----------------
  x$stattheme_data$zhdiagonal <- c("#407B9F", "#3FA5A5", "#6CBD72", "#EDDD79",
                                   "#F4B283", "#F3A5A7", "#E0D6E6", "#857091")


  # zhBlueYellowGrey
  #-----------------
  x$stattheme_data$zhbyg <- c( "#003B5E", "#006EB1", "#5FA9D5", "#FFE16F",
                               "#C2D37B", "#86CEB8", "#56BFB9", "#9DB99E",
                               "#656564")

  x$stattheme_data$zhbyglight <- c("#2A5A78", "#2A85BE", "#79B7DC", "#FFE586",
                                   "#CCD991", "#99D6C3", "#72C9C4", "#ACC4AD",
                                   "#7E7E7D")


  # zhlake
  #-------
  x$stattheme_data$zhlake <- c("#002338", "#004671", "#0076BD", "#3F94BA",
                               "#7FB2B7", "#BFD0B4", "#DAEBCC", "#FFEFB2")


  # zhlikert Benchmarking mit Kantonsrot / Kantonsgr?n - '#e2001a', "#3ea743"
  #--------------------------------------------------------------------------
  x$stattheme_data$zhlikert5gr <- c( "#F09B7A", "#E93F53", "#6EBD72", "#B3DA8A",
                                     "#F8F7A2")

  x$stattheme_data$zhlikert5br <- c("#F09B7A", "#E93F53", "#3F98CD", "#9BC7B7",
                                    "#F8F7A2")

  x$stattheme_data$zhlikert6gr <- c("#6EBD72", "#A5D485", "#DCEB98", "#f2a975",
                                    "#EF8872", "#E93F53")

  x$stattheme_data$zhlikert6br <- c("#3F98CD", "#7FB9DD", "#9FDADA", "#F7C5A1",
                                    "#EF8872", "#E93F53")


  # BENCHMARKING-PALETTEN
  #----------------------
  x$stattheme_data$bmlikert <- c("#e2001a", "#f27644", "#fab872", "#ffe16f",
                                 "#9ec359", "#3ea743")

  x$stattheme_data$bmtabelle <- c("#ff9faa", "#f9bda5", "#d6e6b8", "#ffe16f",
                                  "#b9e5bb")

  x$stattheme_data$bmmittelwert <- c("#6DB0D9")


  # 2. Definiere Liste mit sequentiellen Paletten bzw. low / high Farben

  #Da zhdarker nicht dunkel genug ist - mische diese mit schwarz
  #> interpolate(colors$zhdarker, color="#000000",degree=3,7)
  #[1] "#2A526A" "#2A6060" "#3F6141" "#766D39" "#794C2D" "#763132" "#584A5F"

  x$stattheme_data$zhbygseq <- x$stattheme_data$zhbyg[c(4,7,1)]

  x$stattheme_data$zhseq <- c(x$stattheme_data$zhultralight[3],
                              x$stattheme_data$zh[2],
                              x$stattheme_data$zhdarker[1],
                              "#2A526A")

  x$stattheme_data$zhlakeseq <- c(x$stattheme_data$zhlake[8],
                                  x$stattheme_data$zhlake[3],
                                  x$stattheme_data$zhlakelight[1])


  x$stattheme_data$sequential$zhblueseq$low <- "#BFDCED"

  x$stattheme_data$sequential$zhblueseq$high <- "#19242D"


  x$stattheme_data$sequential$zhredseq$low  <- "#F6C5CB"

  x$stattheme_data$sequential$zhredseq$high <-"#983844"


  x$stattheme_data$sequential$zhvioletseq$high <- "#715F7A"

  x$stattheme_data$sequential$zhvioletseq$low  <- "#E3D9E6"


  x$stattheme_data$sequential$zhbgseq$high <- "#0076BD"

  x$stattheme_data$sequential$zhbgseq$low  <- "#B1DBB3"


  x$stattheme_data$sequential$zhgreenseq$high <- "#527F54"

  x$stattheme_data$sequential$zhgreenseq$low  <- "#D2E9D3"


  #3. Definiere Liste mit divergierenden Paletten

  #zhmh - multihue

  x$stattheme_data$diverging$zhmhblue$low  <- "#F3EBAE"

  x$stattheme_data$diverging$zhmhblue$mid  <- "#0076BD"

  x$stattheme_data$diverging$zhmhblue$high <- "#19242D"


  x$stattheme_data$diverging$zhmhgreen$low  <- "#F3EBAE"

  x$stattheme_data$diverging$zhmhgreen$mid  <- "#00a1a3"

  x$stattheme_data$diverging$zhmhgreen$high <- "#19242D"


  x$stattheme_data$diverging$zhbluered$low  <- "#e2001a"

  x$stattheme_data$diverging$zhbluered$mid  <- "#f0f0f0"

  x$stattheme_data$diverging$zhbluered$high <- "#0076BD"


  x$stattheme_data$diverging$zhvioletgreen$high <- "#3ea743"

  x$stattheme_data$diverging$zhvioletgreen$mid  <- "#f0f0f0"

  x$stattheme_data$diverging$zhvioletgreen$low  <- "#885ea0"


  x$stattheme_data$diverging$zhspectral$low  <- "#EA5251"

  x$stattheme_data$diverging$zhspectral$mid  <- "#FFE57F"

  x$stattheme_data$diverging$zhspectral$high <- "#3F98CC"
}


#' stattheme_pal()
#' @description Stat ZH scales

#' @importFrom scales manual_pal

#' @keywords stattheme_pal

#' @noRd
stattheme_pal <- function(palette = "default"){

  if (palette %in% names(x$stattheme_data)){
    scales::manual_pal(unname( x$stattheme_data[[palette]]))

  } else {
    stop(sprintf("palette %s not a valid statZH palette.", palette))
  }
}


#' zhpal
#' Data used by the stat zh palettes
#'
#' @format A \code{list}.
#' @export
zhpal <- x$stattheme_data
