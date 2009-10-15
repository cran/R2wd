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
\details{
%%  ~~ If necessary, more details than the description above ~~
}
\value{
%%  ~Describe the value returned
%%  If it is a LIST, use
%%  \item{comp1 }{Description of 'comp1'}
%%  \item{comp2 }{Description of 'comp2'}
%% ...
}
\references{
%% ~put references to the literature/web site here ~
}
\author{
Christian Ritter
}
\note{
%%  ~~further notes~~
}

%% ~Make other sections like Warning with \section{Warning }{....} ~

\seealso{
%% ~~objects to See Also as \code{\link{help}}, ~~~
}
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
