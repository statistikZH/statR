#' Check if image is an implemented plot type
#'
#' @param image Any R object
#' @keywords internal
#' @noRd
checkImplementedPlotType <- function(image) {

  valid_objects <- c("gg", "ggplot", "histogram", "character")

  correct_object_type <- inherits(image, valid_objects)

  if (!correct_object_type) {
    return(FALSE)

  } else if (is.character(image)) {

    if (!(grepl("(.jpeg|.jpg|.png|.bmp)$", image))) {
      return(FALSE)
    }

    if (!file.exists(image)) {
      warning("Image ", image, " not found!")
    }
  }

  return(TRUE)
}
