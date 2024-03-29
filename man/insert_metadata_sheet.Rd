% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/insert_metadata_sheet.R
\name{insert_metadata_sheet}
\alias{insert_metadata_sheet}
\title{Insert sheet with metadata}
\usage{
insert_metadata_sheet(
  wb,
  sheetname,
  meta_infos,
  logo = getOption("statR_logo"),
  contactdetails = inputHelperContactInfo(compact = TRUE),
  homepage = getOption("statR_homepage"),
  author = "user"
)
}
\arguments{
\item{wb}{workbook object to add new worksheet to.}

\item{sheetname}{Names of the worksheet in output file. Note that this name
will be truncated to 31 characters, must be unique, and cannot contain
some special characters (namely the following: /, \, ?, *, :, [, ]).}

\item{meta_infos}{a list with title, source, and metadata as named objects}

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
}
\description{
Function to add a formatted worksheet with metadata to an
  existing Workbook object.
}
\seealso{
createWorkbook, addWorksheet, writeData
}
\keyword{insert_metadata_sheet}
