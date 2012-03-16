\name{wdQuit}
\Rdversion{1.1}
\alias{wdQuit}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Close Word and remove the handle.
}
\description{
Close Word and remove the handle.
}
\usage{
wdQuit(wdapp = .R2wd)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{wdapp}{
    the handle to the Word Application.
}
}
\details{
  This closes Word. Word may ask back whether you wish to save the graph.
}

%% ~Make other sections like Warning with \section{Warning }{....} ~

\examples{
\dontrun{
wdGet()
wdQuit()
}

}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{IO}
\keyword{connection}

