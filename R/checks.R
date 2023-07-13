#' checkGroupOptionCompatibility()
#'
#' @description Check if group_names and grouplines match
#' @param group_names A character vector with group names
#' @param grouplines A character or integer vector of columns to put grouplines at
#' @keywords internal
#' @noRd
checkGroupOptionCompatibility <- function (group_names, grouplines){
  if (any(!is.na(group_names)) & all(is.na(grouplines))) {
    stop("For a second header, the grouplines must be specified")
  }
}

#' checkImplementedPlotType()
#'
#' @description Check if image is an implemented plot type
#' @param image Any R object
#' @keywords internal
#' @noRd
checkImplementedPlotType <- function(image) {
  length(setdiff(class(image), c("gg", "ggplot", "histogram"))) == 0
}
