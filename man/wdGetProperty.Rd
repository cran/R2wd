\name{wdGetProperty}
\alias{wdGetProperty}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Get a property from Word.
}
\description{
This function descends down a property tree given in the parameter
vector \code{property} and returns the lowest level.
}
\usage{
wdGetProperty(property,object = wdapp[["Selection"]], wdapp = .R2wd)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{property}{
The name of a property or a vector indicating a path in a property tree.
}
  \item{object}{
The object on which the property tree should start.
}
  \item{wdapp}{
The handle to Word.
}
}
\details{
This function descends down the indicated property tree and returns the
lowest level. It can also work with calls to the 'Items' method. In this
case, the corresponding item numbers should be inserted in the
\code{property} vector.
}
\value{
The value of the lowest level of the property tree.
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
\code{\link{wdSetProperty}}
%% ~~objects to See Also as \code{\link{help}}, ~~~
}
\examples{
\dontrun{
wdGet()
wdTitle("Title",paragraph=FALSE)
wdGetProperty(c("Font","Name"))
wdGetProperty(c("Font","Size"))
wdParagraph()
wdWrite("This is in normal font")
wdGetProperty(c("Font","Name"))
wdGetProperty(c("Font","Size"))
}
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{IO}
\keyword{connection}
