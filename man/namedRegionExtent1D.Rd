% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/helper.R
\name{namedRegionRowExtent}
\alias{namedRegionRowExtent}
\alias{namedRegionColumnExtent}
\title{Determine the extent of a named region in a particular direction}
\usage{
namedRegionRowExtent(wb, sheetname, region_name = NULL)

namedRegionColumnExtent(wb, sheetname, region_name = NULL)
}
\arguments{
\item{wb}{A workbook object}

\item{sheetname}{Name of a worksheet}

\item{region_name}{names of regions in Workbook.}
}
\value{
A numeric vector of indices
}
\description{
Get row or column extent of a named region in a workbook object.
}
\keyword{internal}
