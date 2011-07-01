\name{wdStyleSetup}
\Rdversion{1.1}
\alias{wdStyleSetup}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Modify document style
}
\description{
Modify document style
}
\usage{
wdStyleSetup(style = "Normal", fontsize = 11, align=3, wdapp = .R2wd)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{style}{
    default: "Normal"
}
  \item{fontsize}{
    default 11pt
}
  \item{align}{
    type of alignment: 3 corresponds to "justified"
}
  \item{wdapp}{
    handle to the Word Application (usually not needed).
}
}
\details{
  this invokes the required methods in word to set the document style
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
wdStyleSetup()
}
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{IO}
\keyword{connection}
