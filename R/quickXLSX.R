# quickXLSX: Function to create formatted spreadsheet with a single worksheet automatically

#' quickXLSX
#'
#' Function to create a formated single-worksheet XLSX automatically
#' @param data data to be included in the XLSX-table.
#' @param name title of the table / file.
#' @param grouplines insert vertical lines to separate columns visually.
#' @keywords quickXLSX
#' @export
#' @examples
#' quickXLSX(mtcars,"title of the spreadsheet")

# Function

quickXLSX <- function(data, name,grouplines=0) {


  # create workbook
  wb <- openxlsx::createWorkbook(paste(name))

  options("openxlsx.numFmt" = "#,###0")

  # data-container from row 5
  datenbereich = 14

  #define width of the area in which data is contained for formating
  spalten = ncol(data)

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

    ## Add worksheet
  openxlsx::addWorksheet(wb,paste(name))

    # Style ---------------------

    #title,subtitle & Header

    # Titel & Untertitel -----------------

  #Logo
  if (file.exists("L:/STAT/08_DS/06_Diffusion/Logos_Bilder/LOGOS/STAT_LOGOS/Stempel_STAT-01.png")){

    openxlsx::insertImage(wb, 1, "L:/STAT/08_DS/06_Diffusion/Logos_Bilder/LOGOS/STAT_LOGOS/Stempel_STAT-01.png",width=2.145, height=0.7865, units="in")
  }

  #Titel
  openxlsx::writeData(wb, sheet = 1, name, headerStyle=title,startRow = 7)

  ##Quelle
  openxlsx::writeData(wb, sheet = 1, "Quelle: Statistisches Amt des Kantons Zürich", headerStyle=subtitle, startRow = 8)

  ##Quelle
  openxlsx::writeData(wb, sheet = 1, "Bemerkungen:", headerStyle=subtitle, startRow = 9)

  #Kontakt
  openxlsx::writeData(wb, sheet = 1, "Datashop, Tel: 0432597500", headerStyle=subtitle, startRow = 2, startCol=contact)
  openxlsx::writeFormula(wb, 1, x = '=HYPERLINK("mailto:datashop@statistik.ji.zh.ch", "datashop@statistik.ji.zh.ch")'
               ,startRow = 3, startCol=contact)
  openxlsx::writeFormula(wb, 1, x = '=HYPERLINK("http://www.statistik.zh.ch", "www.statistik.zh.ch")'
               ,startRow = 4, startCol=contact)

  # writeData(wb, sheet = i, paste('<a href = "datashop@statistik.ji.zh.ch>'), headerStyle=subtitle, startRow = 3, startCol=spalten-2)

  #Aktualisierungsdatum
  openxlsx::writeData(wb, sheet = 1, paste("Aktualisiert am ", format(Sys.Date(), format="%d.%m.%Y"), " durch: ",
                                           stringr::str_sub(Sys.getenv("USERNAME"),-2)), headerStyle=subtitle, startRow = 5, startCol=contact)

    # writeData(wb, sheet = i, "Wohngeb?ude", headerStyle=subtitle, startRow = datenbereich-1,startCol = 4)

    # writeData(wb, sheet = i, "Andere Geb?ude", headerStyle=subtitle, startRow = datenbereich-1,startCol = 8)
    # Daten

  openxlsx::writeData(wb, sheet = 1, data, rowNames = FALSE, startRow = datenbereich, withFilter = FALSE)

    #Füge Formatierungen ein

  openxlsx::addStyle(wb, sheet = 1, title, rows = 7, cols = 1, gridExpand = TRUE)

  openxlsx::addStyle(wb, sheet = 1, headerline, rows = 5, cols = 1:spalten, gridExpand = TRUE,stack = TRUE)

  openxlsx::addStyle(wb, sheet = 1, header, rows = datenbereich, cols = 1:spalten, gridExpand = TRUE,stack = TRUE)


    # addStyle(wb, sheet = 1, header, rows = datenbereich-1, cols = 1:spalten, gridExpand = TRUE,stack = TRUE)

    if (grouplines!=0){

      openxlsx::addStyle(wb, sheet = 1, leftline, rows = (datenbereich-1):nrow(data), cols = grouplines, gridExpand = TRUE,stack = TRUE)

      # addStyle(wb, sheet = i, bottomline, rows = datenbereich+nrow(t_1k%>% filter(s_j_erhjahr==year))+nrow(t_1g%>% filter(s_j_erhjahr==year)), cols = 1:spalten, gridExpand = TRUE,stack = TRUE)

    }

    #Friere oberste Zeilen ein

  openxlsx::freezePane(wb, sheet=1 ,  firstActiveRow = datenbereich+2)

    # bodyStyle <- createStyle(border="TopBottom", borderColour = "#4F81BD")
    # addStyle(wb, sheet = 1, bodyStyle, rows = 2:6, cols = 1:11, gridExpand = TRUE)
  openxlsx::setColWidths(wb, 1, cols=1:spalten, widths = 20) ## set column width for row names column


  openxlsx::saveWorkbook(wb, paste(name,".xlsx",sep=""), overwrite = TRUE)

}

# TEST

