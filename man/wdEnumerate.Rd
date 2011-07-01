\name{wdEnumerate}
\Rdversion{1.1}
\alias{wdEnumerate}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Start a numbered list
}
\description{
This corresponds to pushing the [123] button in Word.
}
\usage{
wdEnumerate(Gallery=2, Template=1, wdapp = .R2wd)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{Gallery}{
the (numbered) item in the list gallery. The pre-set \code{Gallery=2}
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
Starts a numbered list like pushing on the [123] button in Word.
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
\dontrun{wdEnumerate("R2wd is a package to write MS-Word files from R")}
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{IO}
\keyword{connection}

