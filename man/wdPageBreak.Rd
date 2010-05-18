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
%%  ~~who you are~~
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
