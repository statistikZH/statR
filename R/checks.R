#' checkGroupOptionCompatibility()
#'
#' @description Check if group_names and grouplines match
#' @param group_names A character vector with group names
#' @param grouplines A character or integer vector of columns to put grouplines at
#' @keywords internal
#' @noRd
checkGroupOptionCompatibility <- function(group_names, grouplines){
  if (any(!is.na(group_names)) & all(is.na(grouplines))){
    stop("For a second header, the grouplines must be specified")
  }
}
