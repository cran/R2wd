\name{wdInsertFootnote}
\Rdversion{1.1}
\alias{wdInsertFootnote}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Insert a Footnote
}
\description{
Inserts a footnote at the current selection.
}
\usage{
wdInsertFootnote(text="",reference="",wdapp=.R2wd)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{text}{
the footnote text.
}
  \item{reference}{
the footnote reference. By default, this is generated automatically by Word.
}
  \item{wdapp}{
the handle to the Word Application (usually not needed).
}
}
\details{
%%  ~~ If necessary, more details than the description above ~~
}
\value{
(invisible): a handle to the footnote object. Can be used to manipulate
(usually not used).
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
wdInsertFootnote("see also SWORD (rcom.unive.ac.at) for integrating R
calculations in word documents.")
wdBody("This package allows ...")
}
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{IO}
\keyword{connection}
