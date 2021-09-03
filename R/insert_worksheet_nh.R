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
#' @param grouplines defaults to FALSE. Can be used to separate grouped variables visually.
#' @keywords insert_worksheet
#' @importFrom dplyr "%>%"
#' @noRd
#' @examples
#'
#' # create workbook
#' wb <- openxlsx::createWorkbook("hello")
#'
#' insert_worksheet_nh(mtcars[c(1:10),],wb,"mtcars",c(1:4),"carb",grouplines=c(1,5,6))
#'
#'# create workbook
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
#'\dontrun{
#'  # save workbook
#'  openxlsx::saveWorkbook(export,"example.xlsx")
#'}

# Function

insert_worksheet_nh <- function(data,
                                wb,
                                sheetname="Daten",
                                title="Title",
                                source="statzh",
                                metadata = NA,
                                grouplines = FALSE
) {

  # number of rows filled with metadata
  if(is.na(metadata)){
    n_metadata <- 0
  } else
    n_metadata <- length(metadata)

  datenbereich <- 2 + n_metadata + 3

  # number of data columns
  spalten <- ncol(data)

  # increase width of column names for better auto-fitting of column width
  colnames(data) <- paste0(colnames(data), "  ", sep = "")

  if(nchar(sheetname)>31){
    warning("sheetname is cut to 31 characters (limit imposed by MS-Excel)")
    }

  ## Add worksheet
  openxlsx::addWorksheet(wb,paste(substr(sheetname,0,31)))

  i <- paste(substr(sheetname,0,31))

  # define styles --------------------------------------------------------------
  style_title <- openxlsx::createStyle(fontSize=14, textDecoration="bold",
                                       fontName="Arial")
  style_subtitle <- openxlsx::createStyle(fontSize=11, fontName="Calibri")
  style_header <- openxlsx::createStyle(fontSize = 12, fontName="Calibri",
                                        fontColour = "#000000",  halign = "left",
                                        border="Bottom",  borderColour = "#009ee0",
                                        textDecoration = "bold")
  style_leftline <- openxlsx::createStyle(border="Left", borderColour = "#4F81BD")


  # fill in data ---------------------------------------------------------------
  ## Titel
  openxlsx::addStyle(wb
                     ,sheet = i
                     ,style_title
                     ,rows = 2
                     ,cols = 1)
  openxlsx::writeData(wb
                      ,sheet = i
                      ,title
                      ,headerStyle=style_title
                      ,startRow = 2)

  ## Quelle
  if(source=="statzh"){
    source <- "Quelle: Statistisches Amt des Kantons Zürich"
  }
  openxlsx::addStyle(wb
                     ,sheet = i
                     ,style_subtitle
                     ,rows = 3
                     ,cols = 1
                     ,gridExpand = TRUE
  )
  openxlsx::writeData(wb
                      ,sheet = i
                      ,source
                      ,headerStyle=style_subtitle
                      ,startRow = 3
  )

  ##Metadata
  openxlsx::addStyle(wb
                     ,sheet = i
                     ,style_subtitle
                     ,rows = 4
                     ,cols = 1
                     ,gridExpand = TRUE
  )
  openxlsx::writeData(wb
                      ,sheet = i
                      ,metadata
                      ,headerStyle=style_subtitle
                      ,startRow = 4
  )

  ### Metadaten zusammenmergen
  purrr::walk(1:(4+n_metadata), ~openxlsx::mergeCells(wb, sheet = i, cols = 1:26, rows = .))

  # Daten
  openxlsx::addStyle(wb
                     ,sheet = i
                     ,style_header
                     ,rows = datenbereich
                     ,cols = 1:spalten
                     ,gridExpand = TRUE
                     ,stack = TRUE
  )

  openxlsx::writeData(wb,
                      sheet = i
                      # ,as.data.frame(dpylr::ungroup())
                      ,as.data.frame(dplyr::ungroup(data))
                      ,rowNames = FALSE
                      ,startRow = datenbereich
                      ,withFilter = FALSE
  )

  if (!is.null(grouplines)){
    datenbereich_end <- nrow(data)+datenbereich
    openxlsx::addStyle(wb
                       ,sheet = i
                       ,style_leftline
                       ,rows=datenbereich:datenbereich_end
                       ,cols = grouplines
                       ,gridExpand = TRUE
                       ,stack = TRUE
    )

  }

  # minmale Spaltenbreite definieren
  options("openxlsx.minWidth" = 5)

  # automatische Spaltenbreite
  openxlsx::setColWidths(wb, sheet = i, cols=1:spalten, widths = "auto", ignoreMergedCells = TRUE)
}
