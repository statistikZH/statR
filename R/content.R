#' Helper functions for adding attributes to input objects
#'
#' @description A set of helper functions which serve the purpose of allowing
#'   users to assign information like sheetnames, titles, sources, metadata, and
#'   more to an object. These attributes are then used in place of the function
#'   arguments.
#'
#' @details
#'   \code{add_sheetname}: Adds a sheetname.
#'
#'   \code{add_title}: Adds a title
#'
#'   \code{add_source}: Adds a source. Behavior can be modified further via the
#'     prefix and collapse argument. Setting collapse to a value other than NULL
#'     will result in the input provided in value being concatenated.
#'
#'   \code{add_metadata}: Adds metadata. Behavior can be modified further via the
#'     prefix and collapse argument, just like with \code{add_source}.
#'
#'   \code{add_grouplines}: Adds group lines. Throws an error if input object is
#'     not a data.frame.
#'
#'   \code{add_group_names}: Adds group names. Throws an error if group lines
#'     have not been set.
#'
#'   \code{add_plot_height}: Adds plot height. Throws an error if input object is
#'     not a valid plot type.
#'
#'   \code{add_plot_width}: Adds plot width. Throws an error if input object is
#'     not a valid plot type.
#'
#'   \code{add_plot_size}: Assigns both height and width in one call. Expects a
#'     vector of length 2, with the first element corresponding to width, and the
#'     second to height.
#' @param object The object to add an attribute to
#' @param value A value
#' @param prefix A prefix, default to NULL
#' @param collapse Separator to collapse character vectors on, defaults to NULL
#' @rdname add_attribute
#' @examples
#' library(dplyr)
#' library(ggplot2)
#'
#' # Load data and assign attributes in one pipeline
#' df <- mtcars %>%
#'   add_sheetname("Cars") %>%
#'   add_title("Motor Trend Car Road Tests") %>%
#'   add_source(
#'     c("Henderson and Velleman (1981),",
#'       "Building multiple regression models interactively.",
#'       "Biometrics, 37, 391–411."),
#'     prefix = "Source:",
#'     collapse = " ") %>%
#'   add_metadata(
#'     c("The data was extracted from the 1974 Motor",
#'       "Trend US magazine and comprises fuel consumption",
#'       "and 10 aspects of automobile design and",
#'       "performance for 32 automobiles (1973–74 models)."),
#'     collapse = " ")
#'
#' # Create a plot and assign attributes in one pipeline
#' plt <- (ggplot(mtcars) +
#'   geom_histogram(aes(x = hp))) %>%
#'     add_sheetname("PS") %>%
#'     add_title("Histogram of horsepower") %>%
#'     add_plot_size(c(6, 3))
#'
#' \dontrun{
#' # Generate outputfile using minimal call
#' datasetsXLSX(
#'   file = tempfile(fileext = ".xlsx"),
#'   datasets = list(df, plt))
#' }
#' @export
#'
#'
add_attribute <- function(object, what, value) {
  UseMethod("add_attribute", object)
}

#' @rdname add_attribute
#' @keywords internal
add_attribute.default <- function(object, what, value) {
  attr(object, what) <- value
  class(object) <- c(class(object), "Content")
  return(object)
}

#' @rdname add_attribute
#' @keywords internal
add_attribute.Content <- function(object, what, value) {
  attr(object, what) <- value
  return(object)
}

#' @rdname add_attribute
#' @export
add_sheetname <- function(object, value) {
  add_attribute(object, "sheetname", value)
}

#' @rdname add_attribute
#' @export
add_title <- function(object, value) {
  add_attribute(object, "title", value)
}

#' @rdname add_attribute
#' @export
add_source <- function(object, value) {
  add_attribute(object, "source", value)
}

#' @rdname add_attribute
#' @export
add_metadata <- function(object, value) {
  add_attribute(object, "metadata", value)
}

#' @rdname add_attribute
#' @export
add_grouplines <- function(object, value) {
  if (!inherits(object, "data.frame")) {
    warning("Ignoring attribute: grouplines only relevant for data.frames")
    return(object)
  }

  add_attribute(object, "grouplines", value)
}

#' @rdname add_attribute
#' @export
add_group_names <- function(object, value) {
  if (!inherits(object, "data.frame")) {
    warning("Ignoring attribute: group names only relevant for data.frames")
    return(object)
  }

  if (is.null(extract_attribute(object, "grouplines"))) {
    warning("Ignoring attribute: group names require non-null grouplines")
    return(object)
  }

  add_attribute(object, "group_names", value)
}

#' @rdname add_attribute
#' @export
add_plot_size <- function(object, value) {

  if (all(is.na(value))) return(object)

  if (!checkImplementedPlotType(object)) {
    warning("Tried to attach plot dimension attributes to non-plot object")
    return(object)
  }

  if (length(value) != 2) {
    stop("Expected 2 values but got ", length(value))
  }

  object %>%
    add_attribute("plot_width", value[1]) %>%
    add_attribute("plot_height", value[2])
}

#' Check for presence of keyword attributes in input data.frame
#'
#' A function which checks for presence of certain attributes in an input
#' data.frame.
#' @param object Input data.frame
#' @keywords internal
check_for_attributes <- function(object){
  attributes_to_check <- c("title", "source", "metadata", "grouplines",
                           "group_names", "metadata_sheet", "plot_width",
                           "plot_height")

  check <- names(attributes(object))
  return(any(attributes_to_check %in% check))
}


#' Extract attributes from target object
#'
#' @description Function to extract attributes from objects.
#'   \code{extract_attribute} expects the input object to be the target. When
#'   the target object is nested in a list (as is the case in
#'   \code{datasetsXLSX}), \code{extract_attributes} should be used instead.
#' @param object Object to extract attribute from. In practice a data.frame,
#'   ggplot object, or a character string providing a path to an image.
#' @param object_list A list of objects of arbitrary types.
#' @param which Name of the attribute to extract.
#' @param required_val Boolean, if TRUE tries to look up a default in global
#'   options if attribute not found, and raises an error if none was defined.
#' @rdname extract_attribute
#' @keywords internal
extract_attribute <- function(object, which, required_val = FALSE) {

  value <- attr(object, which)

  if (all(is.null(value))) {

    if (required_val) {
      value <- getOption(paste0("statR_default_", which))

      if (all(is.null(value))) {
        stop("No default value found for required argument ", which)
      }

    } else {
      value <- NA
    }
  }

  return(value)
}
#' @rdname extract_attribute
#' @keywords internal
extract_attributes <- function(object_list, which, required_val = FALSE) {
  values <- list()

  for (i in seq_along(object_list)) {
    values[[i]] <- tryCatch(
      extract_attribute(object_list[[i]], which, required_val),
      error = function(err) {
        stop("missing value in required field '", which, "' of dataset ", i)
      })
  }

  return(values)
}
