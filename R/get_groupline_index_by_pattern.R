#' get_groupline_index_by_pattern()
#'
#' Derive groupline index by matching names
#' @param grouplines ...
#' @param data ...
#' @keywords internal

get_groupline_index_by_pattern <- function(grouplines, data){

  get_lowest_col <- function(groupline, data){
    groupline_numbers_single <- which(grepl(groupline, names(data)))

    out <- min(groupline_numbers_single)

    return(out)
  }

  groupline_numbers <- unlist(lapply(grouplines, function(x){
    get_lowest_col(x, data)
    }))

  return(groupline_numbers)
}
