# Function to create formatted XLSX file with multiple worksheets automatically
#' datasetsXLSX
#'
#' Function to create formatted XLSX files automatically
#' @param file filename of the XLSX file. No Default.
#' @param maintitle main title of first sheet "Inhalt".
#' @param datasets datasets (dataframes or images) to be included in the XLSX-table.
#' @param widths width of figure.
#' @param heights height of figure.
#' @param startrows row coordinate of upper left corner of figure.
#' @param startcols column coordinate of upper left corner of figure.
#' @param sheetnames name of the sheet tab.
#' @param titles title of the table in the worksheet, defaults to "Titel" + the value of the variable used to split the dataset across sheets.
#' @param sources source of the data. Defaults to "statzh".
#' @param metadata1 metadata-information to be included. Defaults to NA.
#' @param auftrag_id order number.
#' @keywords datasetsXLSX
#' @export
#' @examples
#'
#' # example with image / plot
#' # create map
#'
#' plot <- plot(mtcars$mpg)
#'
#' datasetsXLSX(file="datasetsXLSX"
#'             ,maintitle = "nice datasets"
#'             ,datasets = list(head(mtcars),tail(mtcars),plot)
#'             ,sheetnames = c("data1","data2","map")
#'             ,widths = c(0,0,5)
#'             ,heights = c(0,0,2.5)
#'             ,startrows = c(0,0,10)
#'             ,startcols = c(0,0,8)
#'             ,titles = c("Title","Title", "Map")
#'             ,sources = c("Quelle: STATENT", "Quelle: Strukturerhebung", "Quelle: Strukturerhebung")
#'             ,metadata1 = c("Bemerkungen: bla", "Bemerkungen: blabla", "Bemerkungen: blablabla")
#'             ,auftrag_id="A2020_0200"
#')
#'
#'
#' # example without image
#'
#' datasetsXLSX(file="datasetsXLSX2"
#' ,maintitle = "nice datasets"
#' ,datasets = list(head(mtcars),tail(mtcars))
#' ,sheetnames = c("data1","data2")
#' ,titles = c("Title","Title")
#' ,sources = c("Quelle: STATENT", "Quelle: Strukturerhebung")
#' ,metadata1 = c("Bemerkungen: bla", "Bemerkungen: blabla")
#' ,auftrag_id="AS2020_01"
#' )
#'
#'


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
                         sources=NULL,
                         metadata1="",
                         auftrag_id,
                         ...
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
                             widths[i],
                             heights[i],
                             startrows[i],
                             startcols[i])
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

    if(logo=="statzh") logo <- paste0(.libPaths(),"/statR/data/Stempel_STAT-01.png")

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
  contactdetails <- c("Datashop"
                      ,"Tel.:  +41 43 259 75 00",
                      "datashop@statistik.zh.ch")
  openxlsx::writeData(wb
                      ,sheet = "Inhalt"
                      ,contactdetails
                      ,xy = c("O", 2)
  )

  homepage <- "http://www.statistik.zh.ch"
  class(homepage) <- 'hyperlink'
  openxlsx::writeData(wb
                      ,"Inhalt"
                      ,x = homepage
                      ,xy = c("O", 5)
  )

  openinghours <- c("Bürozeiten"
                    ,"Montag bis Freitag"
                    ,"09:00 bis 12:00"
                    ,"13:00 bis 16:00")

  openxlsx::writeData(wb
                      ,sheet = "Inhalt"
                      ,openinghours
                      ,xy = c("R", 2)

  )


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

  # source
  openxlsx::writeData(wb
                      ,"Inhalt"
                      ,"Quelle: Statistisches Amt des Kantons Zürich"
                      ,xy = c("C", 11)
  )

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


