\name{wdItemize}
\Rdversion{1.1}
\alias{wdItemize}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Start a bullet list
}
\description{
This corresponds to pushing the [123] button in Word.
}
\usage{
wdItemize(Gallery=1, Template=1, wdapp = .R2wd)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{Gallery}{
the (numbered) item in the list gallery. The pre-set \code{Gallery=1}
implies that bullets are used.
}
  \item{Template}{
the (numbered) item in the list template. The pre-set \code{Template=1}
implies that the first type is used.
}
  \item{wdapp}{
the handle to word (can be ignored).
}
}
\details{
Starts a bullet list like pushing on the [123] button in Word.
}
\author{
Christian Ritter
}

%% ~Make other sections like Warning with \section{Warning }{....} ~

\examples{
\dontrun{wdItemize("R2wd is a package to write MS-Word files from R")}
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{IO}
\keyword{connection}

