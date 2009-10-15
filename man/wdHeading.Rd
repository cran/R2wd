\name{wdHeading}
\Rdversion{1.1}
\alias{wdHeading}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Add text in Heading style.
}
\description{
Add text in Heading style.
}
\usage{
wdHeading(level = 1, text = "", paragraph = TRUE,  wdapp = .R2wd)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{level}{
    The heading level.
}
  \item{text}{
    The text to write as a heading.
}
  \item{paragraph}{
    Whether a paragraph should start after the heading.
}
  \item{wdapp}{
    The handle to the Word Application (usually not needed).
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
wdHeading(1,"R2wd")
wdHeading(2,"General ",paragraph=FALSE)
wdHeading(2,"Principles")
}}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{IO}
\keyword{connection}

