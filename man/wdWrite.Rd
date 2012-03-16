\name{wdWrite}
\Rdversion{1.1}
\alias{wdWrite}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Write text to word.
}
\description{
Write text in current style and end with a paragraph break if requested.
}
\usage{
wdWrite(text = "", paragraph = FALSE, wdapp = .R2wd)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{text}{
the text to write.
}
  \item{paragraph}{
whether a paragraph break should be added at the end.
}
  \item{wdapp}{
the handle to the Word Application (usually not needed).
}
}
\author{
Christian Ritter
}
\note{
This is almost the same as \code{wdBody}.
}

%% ~Make other sections like Warning with \section{Warning }{....} ~

\examples{
\dontrun{
wdGet()
wdTitle("R2wd: A package to write ...")
wdSetFont(fontname="Symbol")
wdWrite("R2wd is a package for ...")
}
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{IO}
\keyword{connection}

