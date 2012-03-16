\name{wdInsertBookmark}
\Rdversion{1.1}
\alias{wdInsertBookmark}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Insert a Bookmark.
}
\description{
Inserts a Bookmark at the current selection.
}
\usage{
wdInsertBookmark(text, wdapp = .R2wd)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{text}{
the bookmark text.
}
  \item{wdapp}{
the handle to the Word Application (usually not needed).
}
}
\author{
Christian Ritter
}

%% ~Make other sections like Warning with \section{Warning }{....} ~

\examples{
\dontrun{
wdGet()
wdTitle("R2wd a package for writing Word documents from R")
wdInsertBookmark("here")
wdBody("This package allows ...")
wdSection("General")
wdGoToBookmark("here")}
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{IO}
\keyword{connection}
