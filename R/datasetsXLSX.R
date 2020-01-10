# datasetsXLSX: Function to create formatted with multiple worksheets spreadsheets automatically

#' datasetsXLSX
#'
#' Function to create formatted spreadsheets automatically
#' @param datasets data to be included in the XLSX-table.
#' @param file filename of the xlsx-file. No Default.
#' @param name title of the table in the worksheet, defaults to "Titel" + the value of the variable used to split the dataset across sheets.
#' @param sheetvar variable which contains the variable to be used to split the data across several sheets.
#' @param source source of the data. Defaults to "statzh".
#' @param metadata metadata-information to be included. Defaults to NA.
#' @param logo path of the file to be included as logo (png / jpeg / svg). Defaults to "statzh"
#' @param contactdetails contactdetails of the data publisher. Defaults to "statzh".
#' @param grouplines columns to be separated visually by vertical lines.
#' @keywords datasetsXLSX
#' @export
#' @examples
#'  \donttest{
#'
<<<<<<< HEAD:R/datasetsXLSXproto.R
#' # example3
#' datasetsXLSX(file="test3",
#'             datasets = list(head(mtcars),head(diamonds), tail(diamonds)),
#'              sheetnames = c("t1", "t2", "t3"),
#'            titles = c("hi", "hey", "hoi"),
#'             sources = c("ji", "hu", "bu"),
#'             metadata1 = c("gut", "schlecht", "neutral")
#' )
#'
#'
#'
#'

=======
#' # example
#' datasetsXLSX(file="t8",
#'            datasets = c(head(mtcars),head(diamonds)),
#'            sheetnames = c("t1", "t2"),
#'            titles = c("hi", "hey"),
#'             sources = c("ji", "hu"),
#'            metadata1 = c("gut", "schlecht")
#')
#' sql1 <- head(mtcars)
#' sql4 <- head(diamonds)
#'
# example2
#' datasetsXLSX(file="test",
#'             datasets = c(sql1, sql2),
#'            sheetnames = "t1",
#'            titles = "hey",
#'            sources = "hu",
#'            metadata1 = "gut"
#') }
>>>>>>> ea176025b2a9c460497a55032394a7562a8607f3:R/datasetsXLSX.R


library(openxlsx)

datasetsXLSX <- function(file,
<<<<<<< HEAD:R/datasetsXLSXproto.R
                         maintitle,
=======

>>>>>>> ea176025b2a9c460497a55032394a7562a8607f3:R/datasetsXLSX.R
                         datasets,
                         sheetnames,
                         titles,
                         sources,
                         metadata1,
                         auftrag_id,
                         ...){



  wb <- openxlsx::createWorkbook("hello")

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

    #dynamisch mit sheetvar!
    statR::insert_worksheet(data=dataset,
                            workbook=wb,
                            sheetname = sheetnames_def,
                            title = title_def,
                            source = source_def,
                            metadata = metadata_def
                            )
  }


  # Create index sheet
  openxlsx::addWorksheet(wb,"Inhalt")

  #  hide gridlines
  showGridLines(wb
                ,sheet = "Inhalt"
                ,showGridLines = F
  )

  ## set col widths
  setColWidths(wb
               ,"Inhalt"
               ,cols = 1
               ,widths = 1
  )

  ## Insert images
  insertImage(wb
              ,"Inhalt"
              ,"C:/gitrepos/Raum_Siedlungsreport/Data/Stempel_STAT-01.png"
              ,startRow = 2
              ,startCol = 2
              ,width = 2.5
              ,height = 0.9
              ,units = "in"
  )

  # contact
  contactdetails <- c("Datashop"
                      ,"Tel.:  +41 43 259 75 00",
                      "datashop@statistik.zh.ch")
  writeData(wb
            ,sheet = "Inhalt"
            ,contactdetails
            ,xy = c("O", 2)
  )

  homepage <- "http://www.statistik.zh.ch"
  class(homepage) <- 'hyperlink'
  writeData(wb
            ,"Inhalt"
            ,x = homepage
            ,xy = c("O", 5)
  )

  openinghours <- c("Bürozeiten"
                    ,"Montag bis Freitag"
                    ,"09:00 bis 12:00"
                    ,"13:00 bis 16:00")

  writeData(wb
            ,sheet = "Inhalt"
            ,openinghours
            ,xy = c("R", 2)

  )


  # headerline
  headerline <- createStyle(border="Bottom", borderColour = "#009ee0",borderStyle = getOption("openxlsx.borderStyle", "thick"))
  addStyle(wb
           ,"Inhalt"
           ,headerline
           ,rows = 6
           ,cols = 1:20
           ,gridExpand = TRUE
           ,stack = TRUE
  )

  #Erstellungsdatum
  writeData(wb
            ,"Inhalt"
            ,paste("Erstellt am "
                   ,format(Sys.Date(), format="%d.%m.%Y"))
            ,xy = c("O", 8)
  )


  if (!is.null(auftrag_id)){
    # Auftragsnummer
    writeData(wb
              ,sheet = "Inhalt"
              ,paste("Auftragsnr.:", auftrag_id)
              ,xy = c("O", 9)
    )
  }

  # title
  titleStyle <- createStyle(fontSize=20, textDecoration="bold",fontName="Arial", halign = "left")
  addStyle(wb
           ,"Inhalt"
           ,titleStyle
           ,rows = 10
           ,cols = 3
           ,gridExpand = TRUE
           )

  writeData(wb
            ,"Inhalt"
            ,x = maintitle
            ,headerStyle=titleStyle
            ,xy = c("C", 10)
            )

  # source
  writeData(wb
            ,"Inhalt"
            ,"Quelle: Statistisches Amt des Kantons Zürich"
            ,xy = c("C", 11)
            )

  # subtitle
  subtitleStyle <- createStyle(fontSize=11, textDecoration="bold",fontName="Arial", halign = "left")
  addStyle(wb
           ,sheet = "Inhalt"
           ,subtitleStyle
           ,rows = 14
           ,cols = 3
           ,gridExpand = TRUE
           )

  writeData(wb
            ,sheet = "Inhalt"
            ,x = "Inhalt"
            ,headerStyle=subtitleStyle
            ,xy = c("C", 13)
            )

  ## Writing internal hyperlinks
  writeFormula(wb
               ,sheet = "Inhalt"
               ,x = makeHyperlinkString(sheet = sheetnames
                                        #,row = 1
                                        #,col = 1
                                        ,text = titles # soll arguments titles entsprechen
               )
               ,xy = c("C", 15)
               )


  openxlsx::worksheetOrder(wb) <- c(length(names(wb)),1:(length(names(wb))-1))


  #openxlsx::worksheetOrder(wb)<-rev(openxlsx::worksheetOrder(wb))



  openxlsx::saveWorkbook(wb, paste(file, ".xlsx", sep = ""))
}

<<<<<<< HEAD:R/datasetsXLSXproto.R
#
# # example1
datasetsXLSX(file="aloha",
             maintitle = "test",
             datasets = list(head(mtcars),head(mtcars)),
             sheetnames = c("t1", "t2"),
             titles = c("hi", "hey"),
             sources = c("ji", "hu"),
             metadata1 = c("HAE", "schlecht"),
             auftrag_id="AAAAA"
             )


datasetsXLSX(file2="test3",
                       datasets = list(head(mtcars),head(diamonds), tail(diamonds)),
                          sheetnames = c("t1", "t2", "t3"),
                      titles = c("hi", "hey", "hoi"),
                        sources = c("ji", "hu", "bu"),
                        metadata1 = c("gut", "schlecht", "neutral"))
#
# # example2
# datasetsXLSX(file="test2",
#              datasets = list(head(mtcars),head(diamonds)),
#              sheetnames = "t1",
#              titles = "hey",
#              sources = "hu",
#              metadata1 = "gut"
# )
#
# # example3
# datasetsXLSX(file="test3",
#              datasets = list(head(mtcars),head(diamonds), tail(diamonds)),
#              sheetnames = c("t1", "t2", "t3"),
#              titles = c("hi", "hey", "hoi"),
#              sources = c("ji", "hu", "bu"),
#              metadata1 = c("gut", "schlecht", "neutral")
# )

=======
>>>>>>> ea176025b2a9c460497a55032394a7562a8607f3:R/datasetsXLSX.R
