\name{wdPageBreak}
\Rdversion{1.1}
\alias{wdPageBreak}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Insert a page break.
}
\description{
Insert a page break.
}
\usage{
wdPageBreak(wdapp = .R2wd)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{wdapp}{
    the handle to Word Application (usually not needed).
}
}

%% ~Make other sections like Warning with \section{Warning }{....} ~

\examples{
\dontrun{
wdGet()
wdTitle("R2wd: Writing Word Documents from R")
wdBody("R2wd permits writing ...")
wdPageBreak(continuous=FALSE)
wdBody("It goes on on the next page")
}
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{IO}
\keyword{connection}
