# statXLSX: Function to create formatted with multiple worksheets spreadsheets automatically

#' statXLSX
#'
#' Function to create formatted spreadsheets automatically
#' @param data data to be included in the XLSX-table.
#' @param name title of the table / file.
#' @param sheets values of the discrete or categorical variable for which a single worksheet should be generated
#' @param sheetvar variable which contains the discrete variable to be used as 'period'
#' @param grouplines defaults to FALSE. Can be used to separate grouped variables visually.
#' @keywords statXLSX
#' @export
#' @examples
#' # Generation of a spreadsheet with four worksheets (one per 'carb'-category).
#' # Can be used to generate worksheets for multiple years.
#'
#' statXLSX(mtcars[c(1:10),],"mtcars",c(1:4),carb,grouplines=c(1,5,6))

# Function

statXLSX <- function(data, name, sheets,sheetvar, grouplines = FALSE) {

  #Zahlenformat: Tausendertrennzeichen
  options("openxlsx.numFmt" = "#,###0")

  #extrahiere colname
  col_name <- rlang::enquo(sheetvar)

  # create workbook
  wb <- openxlsx::createWorkbook(paste(name))

  # data-container from row 5
  datenbereich = 14

  #define width of the area in which data is contained for formating
  spalten = ncol(data)

  #position of contact details
  contact = if(spalten>4){spalten-2}else{3}

  #styles
  title <- openxlsx::createStyle(fontSize=14, textDecoration="bold",fontName="Arial")

  subtitle <- openxlsx::createStyle(fontSize=12, textDecoration="italic",fontName="Arial")

  header <- openxlsx::createStyle(fontSize = 12, fontColour = "#000000",  halign = "left", border="Bottom",  borderColour = "#009ee0",textDecoration = "bold")

  # header1 <- createStyle(fontSize = 14, fontColour = "#FFFFFF", halign = "center",
  #                        fgFill = "#3F98CC", border="TopBottom", borderColour = "#4F81BD")

  # header2 <- createStyle(fontSize = 12, fontColour = "#FFFFFF", halign = "center",
  #                        fgFill = "#407B9F", border="TopBottom", borderColour = "#4F81BD")

  headerline <- openxlsx::createStyle(border="Bottom", borderColour = "#009ee0",borderStyle = getOption("openxlsx.borderStyle", "thick"))

  #Linien
  bottomline <- openxlsx::createStyle(border="Bottom", borderColour = "#009ee0")

  leftline <- openxlsx::createStyle(border="Left", borderColour = "#4F81BD")


  ### Loop for multiple years / worksheets ------------

  for (year in sheets){

    #get index
    i<- which(sheets==year)

    ## Add worksheet
    openxlsx::addWorksheet(wb,paste(year))

    # img <- system.file("zh2.png", package = "openxlsx")
    #
    # insertImage(wb, 1, img, startRow = 1,  startCol = 1)

    # Style ---------------------

    #title,subtitle & Header


    # Titel & Untertitel -----------------

    #Logo
    if (file.exists("L:/STAT/08_DS/06_Diffusion/Logos_Bilder/LOGOS/STAT_LOGOS/Stempel_STAT-01.png")) {
      openxlsx::insertImage(wb, i, "L:/STAT/08_DS/06_Diffusion/Logos_Bilder/LOGOS/STAT_LOGOS/Stempel_STAT-01.png",width=2.145, height=0.7865, units="in")
    }
    #Titel
    openxlsx::writeData(wb, sheet = i, name, headerStyle=title,startRow = 7)

    ##Quelle
    openxlsx::writeData(wb, sheet = i, "Quelle: Statistisches Amt des Kantons Zürich", headerStyle=subtitle, startRow = 8)

    ##Quelle
    openxlsx::writeData(wb, sheet = i, "Bemerkungen:", headerStyle=subtitle, startRow = 9)

    #Kontakt
    openxlsx::writeData(wb, sheet = i, "Datashop, Tel: 0432597500", headerStyle=subtitle, startRow = 2, startCol=contact)
    openxlsx::writeFormula(wb, i, x = '=HYPERLINK("mailto:datashop@statistik.ji.zh.ch", "datashop@statistik.ji.zh.ch")'
                 ,startRow = 3, startCol=contact)
    openxlsx::writeFormula(wb, i, x = '=HYPERLINK("http://www.statistik.zh.ch", "www.statistik.zh.ch")'
                 ,startRow = 4, startCol=contact)

    #Aktualisierungsdatum
    openxlsx::writeData(wb, sheet = i, paste("Aktualisiert am ", format(Sys.Date(), format="%d.%m.%Y"), " durch: ",
                                             stringr::str_sub(Sys.getenv("USERNAME"),-2)), headerStyle=subtitle, startRow = 5, startCol=contact)

    openxlsx::writeData(wb, sheet = i, as.data.frame(data %>% dplyr::filter((!!col_name) == year)%>%ungroup()), rowNames = FALSE, startRow = datenbereich, withFilter = FALSE)

    #Füge Formatierungen ein

    openxlsx::addStyle(wb, sheet = i, headerline, rows = 5, cols = 1:spalten, gridExpand = TRUE,stack = TRUE)

    openxlsx::addStyle(wb, sheet = i, title, rows = 7, cols = 1, gridExpand = TRUE)

    # addStyle(wb, sheet = i, header1, rows = datenbereich, cols = 1:spalten, gridExpand = TRUE,stack = TRUE)

    openxlsx::addStyle(wb, sheet = i, header, rows = datenbereich, cols = 1:spalten, gridExpand = TRUE,stack = TRUE)

    if (!is.null(grouplines)){

      openxlsx::addStyle(wb, sheet = i, leftline, rows =datenbereich:nrow(data), cols = grouplines, gridExpand = TRUE,stack = TRUE)

    }

    #Friere oberste Zeilen ein

    openxlsx::freezePane(wb, sheet=i ,  firstActiveRow = datenbereich+2)

    # bodyStyle <- createStyle(border="TopBottom", borderColour = "#4F81BD")
    # addStyle(wb, sheet = 1, bodyStyle, rows = 2:6, cols = 1:11, gridExpand = TRUE)
    openxlsx::setColWidths(wb, i, cols=4:spalten, widths = 18) ## set column width for row names column

  }

  openxlsx::worksheetOrder(wb)<-rev(openxlsx::worksheetOrder(wb))


  openxlsx::saveWorkbook(wb, paste(name,".xlsx",sep=""), overwrite = TRUE)

}
