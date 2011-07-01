\name{wdSetProperty}
\alias{wdSetProperty}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Set a property in Word.
}
\description{
This function descends down a property tree given in the parameter
vector \code{property} and sets the lowest level to \code{value}.
}
\usage{
wdSetProperty(property, value, object = wdapp[["Selection"]], wdapp = .R2wd)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{property}{
The name of a property or a vector indicating a path in a property tree.
}
  \item{value}{
The value, the lowest level in the property tree should be set to.
}
  \item{object}{
The object on which the property tree should start.
}
  \item{wdapp}{
The handle to Word.
}
}
\details{
This function descends down the indicated property tree. It can also
work with calls to the 'Items' method. In this case, the corresponding
item numbers should be inserted in the \code{property} vector.
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
\code{\link{wdGetProperty}}
%% ~~objects to See Also as \code{\link{help}}, ~~~
}
\examples{
\dontrun{
wdGet()
wdTitle("Title")
wdWrite("This is in normal font")
wdSetProperty(c("Font","Name"),"Garamond")
wdSetProperty(c("Font","Size"),20)
wdWrite("This is in 14pt Garamond")
wdNormal("This switches back to Normal")
}
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{IO}
\keyword{connection}

