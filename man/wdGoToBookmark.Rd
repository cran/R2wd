\name{wdGoToBookmark}
\Rdversion{1.1}
\alias{wdGoToBookmark}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Navigate to a bookmark in the active Word document.
}
\description{
Looks up the \code{bookmark} in the bookmarks collection and goes to it.
}
\usage{
wdGoToBookmark(bookmark, wdapp = .R2wd)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{bookmark}{
    the name of the bookmark.
}
  \item{wdapp}{
    the handle to the Word application (usually not needed).
}
}

%% ~Make other sections like Warning with \section{Warning }{....} ~

\examples{
\dontrun{
wdGet()
## the following command starts a new section and adds a bookmark.
wdSection("This a new section",label="sec1")
## now we add some text
wdBody("R2wd is a package to write Word documents from R")
## we can now go back to the section header
wdGoToBookmark("sec1")
}
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{IO}
\keyword{connection}

