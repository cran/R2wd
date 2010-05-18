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
wdWrite(text = "", paragraph = TRUE, wdapp = .R2wd)
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
This is almost the same as \code{wdBody}.
}

%% ~Make other sections like Warning with \section{Warning }{....} ~

\seealso{
%% ~~objects to See Also as \code{\link{help}}, ~~~
}
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

