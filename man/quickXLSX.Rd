% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/output_xlsx.R
\name{quickXLSX}
\alias{quickXLSX}
\title{Export a single dataset to a single formatted worksheet}
\usage{
quickXLSX(
  data,
  file,
  title = NULL,
  source = NULL,
  metadata = NULL,
  grouplines = NULL,
  group_names = NULL,
  logo = NA,
  contactdetails = NA,
  homepage = NA,
  author = "user",
  config = "default"
)
}
\arguments{
\item{data}{data to be included.}

\item{file}{file name of the xlsx-file. The extension ".xlsx" is added}

\item{title}{title to be put above the data.}

\item{source}{source of the data. Default can be adjusted via user profiles}

\item{metadata}{metadata information to be included. Defaults to NA, meaning
no metadata are attached.}

\item{grouplines}{Can be used to visually group variables. Values should
either correspond to numeric column indices or column names, denoting the
first variable in a group. Defaults to NA, meaning no lines are added.}

\item{group_names}{A vector of names for groups to be displayed in a
secondary header. Should be of the same length as grouplines, and cannot
be used unless these are set. Defaults to NA, meaning no secondary header
is created.}

\item{logo}{File path to the logo to be included in the index-sheet.
Defaults to the logo of the Statistical Office of Kanton Zurich. This can
either be overridden with a path to an image file, or configured in a user
profile.}

\item{contactdetails}{Character vector with contact information to be displayed
on the title sheet. By default uses \code{inputHelperContactInfo()} to
construct it based on values defined in the active user configuration.}

\item{homepage}{Homepage of data publisher. Default can be adjusted via user
configuration.}

\item{author}{defaults to the last two letters (initials) or numbers of the
internal user name.}

\item{config}{which config file should be used. Default: default}
}
\description{
A simple function for exporting data from R to a single formatted
  .xlsx-spreadsheet.
}
\examples{

title <- "Motor trend car road tests"
source <- paste(
  "Henderson and Velleman (1981). Building",
  "multiple regression models interactively.",
  "Biometrics, 37, 391–411.")
metadata <- paste(
  "The data was extracted from the 1974 Motor",
  "Trend US magazine and comprises fuel consumption",
  "and 10 aspects of automobile design and",
  "performance for 32 automobiles (1973–74 models).")
\dontrun{
quickXLSX(file = tempfile(fileext = ".xlsx"),
          data = mtcars,
          title = title,
          source = source,
          metadata = metadata)
}
}
\keyword{quickXLSX}
