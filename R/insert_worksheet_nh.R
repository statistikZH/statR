#' insert_worksheet_nh
#'
#' Function to add formatted worksheets without headers to an existing workbook
#'
#' @param data data to be included in the XLSX-table.
#' @param wb workbook object to write new worksheet in.
#' @param title title of the table and the sheet
#' @param sheetname name of the sheet-tab.
#' @param source source of the data. Defaults to "statzh".
#' @param metadata metadata-information to be included. Defaults to NA.
#' @param grouplines defaults to NA. Can be used to separate grouped variables visually.
#' @param group_names Name(s) of the second header(s). Format: List e.g list(c("title 1", "title 2", "title 3"))
#' @keywords insert_worksheet
#' @export
#' @importFrom dplyr "%>%"
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
#'                   ,source = "Quelle: Statistisches Amt Kanton Z\u00fcrich"
#'                   ,metadata = "Bemerkung: ...")
#'
#'  # insert a further worksheet
#' insert_worksheet_nh(tail(mtcars)
#'                    ,export
#'                    ,"data2"
#'                    ,title = "Title"
#'                    ,source = "Quelle: Statistisches Amt Kanton Z\u00fcrich"
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
                                grouplines = NA,
                                group_names = NA
){

  # number of data columns
  spalten <- ncol(data)

  # increase width of column names for better auto-fitting of column width
  colnames(data) <- paste0(colnames(data), "  ", sep = "")

  # Check length of sheetname
  sheetname <- check_sheetname(sheetname)

  ## Add worksheet
  openxlsx::addWorksheet(wb, sheetname)


  # fill in data ---------------------------------------------------------------
  ## Titel
  openxlsx::addStyle(wb,
                     sheet = sheetname,
                     style = style_title(),
                     rows = 1,
                     cols = 1)
  openxlsx::writeData(wb,
                      sheet = sheetname,
                      x = title,
                      headerStyle=style_title(),
                      startRow = 1)

  ## Quelle
  sources_to_insert <- prep_source(source)
  source_end_row <- add_additional_text(wb, sources_to_insert, sheetname, 2)


  ## Metadata
  if (!is.na(metadata)){
    metadata_to_insert <- prep_metadata(metadata)
    metadata_start_row <- source_end_row + 1

    metadata_end_row <- add_additional_text(wb, metadata_to_insert, sheetname, metadata_start_row)

    data_start_row <- metadata_end_row + 2
    merge_end_row <- metadata_end_row

  } else{
    data_start_row <- source_end_row + 2
    merge_end_row <- source_end_row
  }

  ### Metadaten zusammenmergen
  purrr::walk(1:merge_end_row,
              ~openxlsx::mergeCells(wb,
                                    sheet = sheetname,
                                    cols = 1:26,
                                    rows = .))

  # Zweite header Zeile einfügen
  if (any(is.null(group_names))){
    group_names <- NA
  }

  if (!any(is.na(group_names))){
    insert_second_header(wb, sheetname, data_start_row, group_names, grouplines, data)
    data_start_row <- data_start_row + 1
  }

  # Daten
  openxlsx::addStyle(wb,
                     sheet = sheetname,
                     style = style_header(),
                     rows = data_start_row,
                     cols = 1:spalten,
                     gridExpand = TRUE,
                     stack = TRUE)

  openxlsx::writeData(wb,
                      sheet = sheetname,
                      x = as.data.frame(dplyr::ungroup(data)),
                      rowNames = FALSE,
                      startRow = data_start_row,
                      withFilter = FALSE)

  if(is.null(grouplines)){
    grouplines <- NA
  }

  if (any(!is.na(grouplines))){
    if (!any(is.na(group_names))){
      data_start_row <- data_start_row - 1
      data_end_row <- nrow(data)+data_start_row +1

    } else {
      data_end_row <- nrow(data)+data_start_row
    }



    if (is.numeric(grouplines)){
      groupline_numbers <- grouplines

    } else if(is.character(grouplines)){
      groupline_numbers <- get_groupline_index_by_pattern(grouplines, data)
    }

    openxlsx::addStyle(wb,
                       sheet = sheetname,
                       style = style_leftline(),
                       rows=data_start_row:data_end_row,
                       cols = groupline_numbers,
                       gridExpand = TRUE,
                       stack = TRUE)
  }

  # minmale Spaltenbreite definieren
  options("openxlsx.minWidth" = 5)

  # automatische Spaltenbreite
  openxlsx::setColWidths(wb, sheet = sheetname, cols=1:spalten, widths = "auto", ignoreMergedCells = TRUE)
}
