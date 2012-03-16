\name{wdPageSetup}
\Rdversion{1.1}
\alias{wdPageSetup}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Modify page setup in Word.
}
\description{
Modify page setup in Word.
}
\usage{
wdPageSetup(orientation = "portrait", margins = rep(1, 4), scope = "section", wdapp = .R2wd)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{orientation}{
    page orientation, either "portrait" or "landscape"
}
  \item{margins}{
    a vector of page margins in inches (a unit variable will be supplied later)
}
  \item{scope}{
    whether the change applies to the entire document ("all") or only to
    the present section ("section").
}
  \item{wdapp}{
    handle to the Word Application (usually not needed).
}
}
\details{
  this invokes the required methods in word to set page orientation and margins
}

%% ~Make other sections like Warning with \section{Warning }{....} ~

\examples{
\dontrun{
 wdGet()
wdPageSetup(orientation="landscape",margins=c(2,2,2,2),scope="all")
}
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{IO}
\keyword{connection}
