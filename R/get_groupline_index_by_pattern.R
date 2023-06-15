#' get_groupline_index_by_pattern()
#'
#' @description Derive groupline index by matching names
#' @inheritParams insert_worksheet
#' @keywords internal
get_groupline_index_by_pattern <- function(grouplines, data){

  get_lowest_col <- function(groupline, data){
    min(which(grepl(groupline, names(data))))
  }

  unlist(lapply(grouplines, function(x) get_lowest_col(x, data)))
}
