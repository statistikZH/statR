
#' isImplementedPlotType()
#'
#' @description Check if object is a supported plot object
#' @returns A logical value
#' @keywords internal
#'
isImplementedPlotType <- function(object) {
  implemented_plot_types <- c("gg", "ggplot", "histogram", "character")

  return(length(setdiff(class(object), implemented_plot_types)) == 0)
}

#' addContentObject()
#'
#' @description Add an object, either a data.frame or a supported plot object
#'   to a list together with
#' @inheritParams namedRegionExtent
#' @returns A numeric vector of row indices
#' @export
#'
addContentObject <- function(list = list(), object, sheetname, title, sources,
                             metadata = NA, grouplines = NA, group_names = NA,
                             width = 5, height = 5) {

  if (isImplementedPlotType(object)) {

    if (!("data" %in% names(list))) list[["data"]] <- list()
    list[["data"]] <- addDataObject(list[["data"]], object, sheetname, title,
                                    sources, metadata, grouplines, group_names)

  } else if (is.data.frame(object)) {

    if (!("plot" %in% names(list))) list[["plot"]] <- list()
    list[["data"]] <- addPlotObject(list[["data"]], object, sheetname,
                                    width, height)
  }
  list <- `attr<-`(list, "configured", TRUE)
}

#' namedRegionRowExtent()
#'
#' @description Get row extent of a named region in a workbook object
#' @inheritParams addContentObject
#' @returns A numeric vector of row indices
#' @keywords internal
#'
addPlotObject <- function(list, object, sheetname, width, height) {

  list[[length(list) + 1]] <- list(
    object, sheetname, width, height
  )
  return(list)
}

#' addDataObject()
#'
#' @description Add a data.frame with metadata to a list for use in datasetsXLSX
#' @inheritParams addContentObject
#' @returns A list
#' @keywords internal
#'
addDataObject <- function(list, object, sheetname, title, sources, metadata,
                          grouplines, group_names) {

  list[[length(list) + 1]] <- list(
    object, sheetname, title, sources, metadata, grouplines, group_names
  )
  return(list)
}
