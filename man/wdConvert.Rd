\name{wdConvert}
\Rdversion{1.1}
\alias{wdConvert}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Unit conversion (inches, points, cm)
}
\description{
This function converts between inches, cm, and points.
}
\usage{
wdConvert(input,from="in",to="pt")
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{input}{
    the number to convert
}
  \item{from}{
    either "in","cm", or "pt".
}
  \item{to}{
    either "in","cm", or "pt".
}
}
\details{
Converts the input value from \code{from} units to \code{to}units.
}
\value{
the converted number.
}
\author{
Christian Ritter
}

%% ~Make other sections like Warning with \section{Warning }{....} ~

\examples{
\dontrun{
wdConvert(1,"in","pt")
wdConvert(72,"pt","in")
wdConvert(1,"cm","pt")
## the following gives an error:
wdConvert(1,"cm","mm")
}
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{IO}
\keyword{connection}
