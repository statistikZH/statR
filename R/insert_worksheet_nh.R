# insert_worksheet_nh: add formatted worksheets without header to an existing Workbook

#' insert_worksheet_nh
#'
#' Function to create formatted spreadsheets automatically
#' @param data data to be included in the XLSX-table.
#' @param workbook workbook object to write new worksheet in.
#' @param title title of the table and the sheet
#' @param sheetname name of the sheet-tab.
#' @param source source of the data. Defaults to "statzh".
#' @param metadata metadata-information to be included. Defaults to NA.
#' @param logo path of the file to be included as logo (png / jpeg / svg). Defaults to "statzh"
#' @param contactdetails contactdetails of the data publisher. Defaults to "statzh".
#' @param grouplines defaults to FALSE. Can be used to separate grouped variables visually.
#' @keywords insert_worksheet
#' @importFrom dplyr "%>%"
#' @export
#' @examples
#'  # insertion of two worksheets in a existing workbook
#'
#'   # create workbook
#' wb <- openxlsx::createWorkbook("hello")
#'
#' insert_worksheet_nh(mtcars[c(1:10),],wb,"mtcars",c(1:4),"carb",grouplines=c(1,5,6))
#'
#'   # create workbook
#' export <- openxlsx::createWorkbook("export")
#'
#' # insert a new worksheet
#' insert_worksheet_nh(head(mtcars)
#'                   ,export
#'                   ,"data1"
#'                  ,title = "Title"
#'                   ,source = "Quelle: Statistisches Amt Kanton Zürich"
#'                   ,metadata = "Bemerkung: ...")
#'
#'  # insert a further worksheet
#' insert_worksheet_nh(tail(mtcars)
#'                    ,export
#'                    ,"data2"
#'                    ,title = "Title"
#'                    ,source = "Quelle: Statistisches Amt Kanton Zürich"
#'                    ,metadata = "Bemerkung: ...")
#'
#'  # save workbook
#'  openxlsx::saveWorkbook(export,"example.xlsx")
#'

# Function

insert_worksheet_nh <- function(data,
                                workbook,
                                sheetname="data",
                                title="Title",
                                source="Quelle: Statistisches Amt Kanton Zürich",
                                metadata = NA,
                                #,logo=NULL
                                grouplines = FALSE
                                #,contactdetails="statzh"
) {

  # Metadata
  remarks <- if (is.na(metadata)) {
    "Bemerkungen:"}
  else if (metadata == "HAE") {"Die Zahlen der letzten drei Jahre sind provisorisch."}
  else {metadata}


  #Zahlenformat: Tausendertrennzeichen
  options("openxlsx.numFmt" = "#,###0")

  #extrahiere colname
  # col_name <- rlang::enquo(sheetvar)

  wb <- workbook

  # data-container from row 5
  n_metadata <- length(metadata)

  datenbereich = 2 + n_metadata + 3

  #define width of the area in which data is contained for formating
  spalten = ncol(data)

  # #position of contact details
  # contact = if(spalten>4){spalten-2}else{3}


  # header1 <- createStyle(fontSize = 14, fontColour = "#FFFFFF", halign = "center",
  #                        fgFill = "#3F98CC", border="TopBottom", borderColour = "#4F81BD")

  # header2 <- createStyle(fontSize = 12, fontColour = "#FFFFFF", halign = "center",
  #                        fgFill = "#407B9F", border="TopBottom", borderColour = "#4F81BD")

  # headerline <- openxlsx::createStyle(border="Bottom", borderColour = "#009ee0",borderStyle = getOption("openxlsx.borderStyle", "thick"))

  #Linien
  bottomline <- openxlsx::createStyle(border="Bottom", borderColour = "#009ee0")

  leftline <- openxlsx::createStyle(border="Left", borderColour = "#4F81BD")

  #wrap
  wrap <- openxlsx::createStyle(wrapText = TRUE)


  # warning if sheetname is longer than the limit imposed by excel (31 characters)
  if(nchar(sheetname)>31){warning("sheetname is cut to 31 characters (limit imposed by MS-Excel)")}

  ## Add worksheet
  openxlsx::addWorksheet(wb,paste(substr(sheetname,0,31)))

  i <- paste(substr(sheetname,0,31))


  #styles
  titleStyle <- openxlsx::createStyle(fontSize=14, textDecoration="bold",fontName="Arial")

  subtitle <- openxlsx::createStyle(fontSize=11, fontName="Calibri")

  header <- openxlsx::createStyle(fontSize = 12, fontName="Calibri", fontColour = "#000000",  halign = "left", border="Bottom",  borderColour = "#009ee0",textDecoration = "bold")


  #Titel
  openxlsx::addStyle(wb
                     ,sheet = i
                     ,titleStyle
                     ,rows = 2
                     ,cols = 1
                     #,gridExpand = TRUE
  )
  openxlsx::writeData(wb
                      ,sheet = i
                      ,title
                      ,headerStyle=titleStyle
                      ,startRow = 2
  )


  ##Quelle
  openxlsx::addStyle(wb
                     ,sheet = i
                     ,subtitle
                     ,rows = 3
                     ,cols = 1
                     ,gridExpand = TRUE
  )
  openxlsx::writeData(wb
                      ,sheet = i
                      ,source
                      ,headerStyle=subtitle
                      ,startRow = 3
  )



  ##Metadata
  openxlsx::addStyle(wb
                     ,sheet = i
                     ,subtitle
                     ,rows = 4
                     ,cols = 1
                     ,gridExpand = TRUE
  )
  openxlsx::writeData(wb
                      ,sheet = i
                      ,metadata
                      ,headerStyle=subtitle
                      ,startRow = 4
  )


  # Daten abfüllen
  openxlsx::writeData(wb,
                      sheet = i
                      # ,as.data.frame(dpylr::ungroup())
                      ,as.data.frame(dplyr::ungroup(data))
                      ,rowNames = FALSE
                      ,startRow = datenbereich
                      ,withFilter = FALSE
  )

  #Füge Formatierungen ein

  openxlsx::addStyle(wb
                     ,sheet = i
                     ,header
                     ,rows = datenbereich
                     ,cols = 1:spalten
                     ,gridExpand = TRUE
                     ,stack = TRUE
  )

  if (!is.null(grouplines)){

    datenbereich_end <-nrow(data)+datenbereich

    openxlsx::addStyle(wb
                       ,sheet = i
                       ,leftline
                       ,rows=datenbereich:datenbereich_end
                       ,cols = grouplines
                       ,gridExpand = TRUE
                       ,stack = TRUE
    )

  }




}
