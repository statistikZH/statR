% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/display.R
\name{display}
\alias{display}
\title{display()}
\usage{
display(col, border = "light gray", ...)
}
\arguments{
\item{col}{a vector of hex-code colors to be displayed}

\item{border}{border color}

\item{...}{further arguments that can be passed to the plot()-function}
}
\description{
Preview of colors specified from hex-codes
}
\examples{
#display can be used to preview all the palettes included in the 'zhpal'-list:
display(zhpal$zhdiagonal)
#example with a single hex-code
display("#000000")
}
\keyword{display}
