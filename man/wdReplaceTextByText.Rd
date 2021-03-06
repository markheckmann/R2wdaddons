% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/addons.R
\name{wdReplaceTextByText}
\alias{wdReplaceTextByText}
\title{Searches and replaces a string}
\usage{
wdReplaceTextByText(find = "", text, wdapp = .R2wd)
}
\arguments{
\item{find}{A string.}

\item{text}{Replacement text.}
}
\value{
Invisibly returns a pointer to the range object.
}
\description{
This function replaces a piece of text with another text.
}
\section{TODO}{
 testing
}

\examples{
file <- system.file("inst/template.docx", package = "R2wdaddons")
wdGet2(file)

# replace placeholder text
wdReplaceTextByText("[placeholder 4]", "SOME REPLACEMENT TEXT")
}
