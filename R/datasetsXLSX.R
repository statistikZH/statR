#' datasetsXLSX()
#'
#' Function to export several datasets and/or figures from R to an .xlsx-file. The function creates an overview sheet and separate sheets
#' for each dataset/figure.
#'
#' When including figures, the heights and widths need to be specified as a vector. For example, say you have two datasets and one figure
#' that you would like to export. widths = c(0,0,5) then suggests that the figure will be 5 inches wide and placed in the third (and last) sheet of the file.
#' The same logic applies to the specification of `startrows` and `startcols`.
#'
#' @param file file name of the spreadsheet. The extension ".xlsx" is added automatically.
#' @param maintitle Title to be put on the first (overview) sheet.
#' @param datasets datasets or plots to be included.
#' @param widths width of figure in inch (1 inch = 2.54 cm). See details.
#' @param heights height of figure in inch (1 inch = 2.54 cm). See details.
#' @param startrows row where upper left corner of figure should be placed. See details.
#' @param startcols column where upper left corner of figure should be placed. See details.
#' @param sheetnames names of the sheet tabs.
#' @param titles titles of the different sheets.
#' @param logo file path to the logo to be included in the index-sheet (default: Logo of the Statistical Office ZH)
#' @param titlesource source to be mentioned on the title sheet beneath the title
#' @param sources source of the data. Defaults to "statzh".
#' @param metadata1 metadata information to be included. Defaults to NA.
#' @param auftrag_id order number.
#' @param contact contact information on the title sheet. Defaults to statzh
#' @param homepage web address to be put on the title sheet. Default to statzh
#' @param openinghours openinghours written on the title sheet. Defaults to Data Shop
#' @keywords datasetsXLSX
#' @export
#' @examples
#'\donttest{
#' \dontrun{
#' plot <- plot(mtcars$mpg)
#'
#'# Example with two datasets and no figure
#'dat1 <- mtcars
#'dat2 <- PlantGrowth
#'
#'datasetsXLSX(file="twoDatasets", # '.xlsx' wird automatisch hinzugefügt
#'             maintitle = "Autos und Pflanzen",
#'             datasets = list(dat1, dat2),
#'             sheetnames = c("Autos","Blumen"),
#'             titles = c("mtcars-Datensatz","PlantGrowth-Datensatz"),
#'             sources = c("Source: Henderson and Velleman (1981).
#'             Building multiple regression models interactively. Biometrics, 37, 391–411.",
#'                         "Dobson, A. J. (1983) An Introduction to Statistical
#'                         Modelling. London: Chapman and Hall."),
#'             metadata1 = c("Bemerkungen zum mtcars-Datensatz: x",
#'                           "Bemerkungen zum PlantGrowth-Datensatz: x"),
#'             auftrag_id="A2021_0000")
#'
#'# Example with two datasets and one figure
#'
#'dat1 <- mtcars
#'dat2 <- PlantGrowth
#'fig <- hist(mtcars$disp)
#'
#'datasetsXLSX(file="twoDatasetsandFigure",
#'             maintitle = "Autos und Pflanzen", # '.xlsx' wird automatisch hinzugefügt
#'             datasets = list(dat1, dat2, fig),
#'             widths = c(0,0,5),
#'             heights = c(0,0,5),
#'             startrows = c(0,0,3),
#'             startcols = c(0,0,3),
#'             sheetnames = c("Autos","Blumen", "Histogramm"),
#'             titles = c("mtcars-Datensatz","PlantGrowth-Datensatz", "Histogramm"),
#'             sources = c("Source: Henderson and Velleman (1981).
#'             Building multiple regression models interactively. Biometrics, 37, 391–411.",
#'                         "Source: Dobson, A. J. (1983) An Introduction to
#'                         Statistical Modelling. London: Chapman and Hall."),
#'             metadata1 = c("Bemerkungen zum mtcars-Datensatz: x",
#'                           "Bemerkungen zum PlantGrowth-Datensatz: x"),
#'             auftrag_id="A2021_0000")
#'}
#'}


# function datasetsXLSX

datasetsXLSX <- function(file,
                         maintitle,
                         datasets,
                         widths,
                         heights,
                         startrows,
                         startcols,
                         sheetnames,
                         titles,
                         logo="statzh",
                         titlesource = "statzh",
                         sources=NULL,
                         metadata1="",
                         auftrag_id,
                         contact = "statzh",
                         homepage = "statzh",
                         openinghours = "statzh"
){

  wb <- openxlsx::createWorkbook("data")

  i<-0

  for (dataset in datasets){

    i <- i+1

    sheetnames_def <- if(length(sheetnames)>1) {
      sheetnames[i]
    } else {i
    }

    title_def <- if(length(titles)>1) {
      titles[i]
    } else {titles
    }

    source_def <- if(length(sources)>1) {
      sources[i]
    } else {sources
    }

    metadata_def <- if(length(metadata1)>1) {
      metadata1[i]
    } else {metadata1
    }


    # #dynamisch mit sheetvar!
    # statR::insert_worksheet2(data=dataset,
    #                         workbook=wb,
    #                         sheetname = sheetnames_def,
    #                         title = title_def,
    #                         source = source_def,
    #                         metadata = metadata_def
    # )

    #dynamisch mit sheetvar!
    if(is.data.frame(dataset)){
      insert_worksheet_nh(data=dataset
                        ,workbook=wb
                        ,sheetname = sheetnames_def
                        ,title = title_def
                        ,source = source_def
                        ,metadata = metadata_def
      )
    }else{ # image
      insert_worksheet_image(dataset,
                             wb,
                             sheetnames_def,
                             width = widths[i],
                             height = heights[i],
                             startrow = startrows[i],
                             startcol = startcols[i])
    }

  }


  # Create index sheet
  openxlsx::addWorksheet(wb,"Inhalt")

  #  hide gridlines
  openxlsx::showGridLines(wb
                          ,sheet = "Inhalt"
                          ,showGridLines = F
  )

  # set col widths
  openxlsx::setColWidths(wb
                         ,"Inhalt"
                         ,cols = 1
                         ,widths = 1
  )

  # insert logo

  if(!is.null(logo)){

    if(logo=="statzh") logo <- paste0(.libPaths(),"/statR/extdata/Stempel_STAT-01.png")

    #
    logo <- logo[file.exists(paste0(.libPaths(),"/statR/extdata/Stempel_STAT-01.png"))]

  openxlsx::insertImage(wb,
                        "Inhalt",
                        file=logo,
                        startRow = 2,
                        startCol = 2,
                        width = 2.5,
                        height = 0.9,
                        units = "in"
  )
  }

  # contact

  if(!is.null(contact)){
    if(any(grepl(contact, pattern = "statzh"))) {
      contact <- c("Datashop"
                   ,"Tel.:  +41 43 259 75 00",
                   "datashop@statistik.zh.ch")
      openxlsx::writeData(wb
                          ,sheet = "Inhalt"
                          ,contact
                          ,xy = c("O", 2)
      )
    } else {
      openxlsx::writeData(wb
                          ,sheet = "Inhalt"
                          ,contact
                          ,xy = c("O", 2))
    }

  }


  if(!is.null(homepage)){

    if(any(grepl(homepage, pattern = "statzh"))) {

      homepage <- "http://www.statistik.zh.ch"
      class(homepage) <- 'hyperlink'
      openxlsx::writeData(wb
                          ,"Inhalt"
                          ,x = homepage
                          ,xy = c("O", 5))
    } else {
      class(homepage) <- 'hyperlink'
      openxlsx::writeData(wb
                          ,"Inhalt"
                          ,x = homepage
                          ,xy = c("O", 5))
    }

  }



  if(!is.null(openinghours)){

    if(any(grepl(openinghours, pattern = "statzh"))) {

      openinghours <- c("Bürozeiten"
                        ,"Montag bis Freitag"
                        ,"09:00 bis 12:00"
                        ,"13:00 bis 16:00")

      openxlsx::writeData(wb
                          ,sheet = "Inhalt"
                          ,openinghours
                          ,xy = c("R", 2)

      )

    } else {

      openxlsx::writeData(wb
                          ,sheet = "Inhalt"
                          ,openinghours
                          ,xy = c("R", 2))

    }

  }


  # headerline
  headerline <- openxlsx::createStyle(border="Bottom", borderColour = "#009ee0",borderStyle = getOption("openxlsx.borderStyle", "thick"))
  openxlsx::addStyle(wb
                     ,"Inhalt"
                     ,headerline
                     ,rows = 6
                     ,cols = 1:20
                     ,gridExpand = TRUE
                     ,stack = TRUE
  )

  #Erstellungsdatum
  openxlsx::writeData(wb
                      ,"Inhalt"
                      ,paste("Erstellt am "
                             ,format(Sys.Date(), format="%d.%m.%Y"))
                      ,xy = c("O", 8)
  )


  if (!is.null(auftrag_id)){
    # Auftragsnummer
    openxlsx::writeData(wb
                        ,sheet = "Inhalt"
                        ,paste("Auftragsnr.:", auftrag_id)
                        ,xy = c("O", 9)
    )
  }

  # title
  titleStyle <- openxlsx::createStyle(fontSize=20, textDecoration="bold",fontName="Arial", halign = "left")
  openxlsx::addStyle(wb
                     ,"Inhalt"
                     ,titleStyle
                     ,rows = 10
                     ,cols = 3
                     ,gridExpand = TRUE
  )

  openxlsx::writeData(wb
                      ,"Inhalt"
                      ,x = maintitle
                      ,headerStyle=titleStyle
                      ,xy = c("C", 10)
  )


  if(any(grepl(titlesource, pattern = "statzh"))){

    # source
    openxlsx::writeData(wb
                        ,"Inhalt"
                        ,"Quelle: Statistisches Amt des Kantons Zürich"
                        ,xy = c("C", 11)
    )

  }else {

    openxlsx::writeData(wb
                        ,"Inhalt"
                        , titlesource
                        ,xy = c("C", 11))

  }




  # subtitle
  subtitleStyle <- openxlsx::createStyle(fontSize=11, textDecoration="bold",fontName="Arial", halign = "left")
  openxlsx::addStyle(wb
                     ,sheet = "Inhalt"
                     ,subtitleStyle
                     ,rows = 14
                     ,cols = 3
                     ,gridExpand = TRUE
  )

  openxlsx::writeData(wb
                      ,sheet = "Inhalt"
                      ,x = "Inhalt"
                      ,headerStyle=subtitleStyle
                      ,xy = c("C", 13)
  )

  ## writing internal hyperlinks
  openxlsx::writeFormula(wb
                         ,sheet = "Inhalt"
                         ,x = openxlsx::makeHyperlinkString(sheet = sheetnames
                                                            #,row = 1
                                                            #,col = 1
                                                            ,text = titles # soll arguments titles entsprechen
                         )
                         ,xy = c("C", 15)
  )


  openxlsx::worksheetOrder(wb) <- c(length(names(wb)),1:(length(names(wb))-1))


  if(missing(metadata1)) {
    metadata1 <- ""
  }


  openxlsx::saveWorkbook(wb, paste(file, ".xlsx", sep = ""))

}



# # # example
# pacman::p_load(tidyverse,tmap)
# data("World")
# # create map
# map <- tm_shape(World) +
#   tm_polygons("HPI")
#
# datasetsXLSX(file="datasetsXLSX"
#              ,maintitle = "nice datasets"
#              ,datasets = list(head(mtcars),tail(mtcars),map)
#              ,sheetnames = c("data1","data2","map")
#              ,widths = c(0,0,10)
#              ,heights = c(0,0,8.5)
#              ,startrows = c(0,0,10)
#              ,startcols = c(0,0,8)
#              ,titles = c("Title","Title", "Map")
#              ,sources = c("Quelle: STATENT", "Quelle: Strukturerhebung", "Quelle: Strukturerhebung")
#              ,metadata1 = c("Bemerkungen: bla", "Bemerkungen: blabla", "Bemerkungen: blablabla")
#              ,auftrag_id="A2020_0200"
# )


