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
#' # Generation of a spreadsheet with four worksheets (one per 'carb'-category).
#' # Can be used to generate worksheets for multiple years.
#'
#' datasetsXLSX(file="test",datasets=list(head(mtcars),head(diamonds)))
#'s


# Functions

sql1 <- head(mtcars)
sql4 <- head(diamonds)

datasetsXLSX <- function(file="test",
                         #datasets1=list(head(mtcars),head(diamonds)),
                         datasets,
                         sheetnames,
                         titles,
                         sources,
                         metadata1,
                         ...){



  wb <- openxlsx::createWorkbook("hello")
  datasets <- list(as.data.frame(datasets[1]), as.data.frame(datasets[2]))

  i<-0

  for (dataset in datasets){

    i <- i+1

    #next 31
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
  openxlsx::saveWorkbook(wb, paste(file, ".xlsx", sep = ""))
}


# example1
datasetsXLSX(file="t8",
             datasets = c(head(mtcars),head(diamonds)),
             sheetnames = c("t1", "t2"),
             titles = c("hi", "hey"),
             sources = c("ji", "hu"),
             metadata1 = c("gut", "schlecht")
             )

# example2
datasetsXLSX(file="test",
             datasets = c(sql1, sql2),
             sheetnames = "t1",
             titles = "hey",
             sources = "hu",
             metadata1 = "gut"
)
