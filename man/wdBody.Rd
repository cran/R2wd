\name{wdBody}
\Rdversion{1.1}
\alias{wdBody}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Write in body text.
}
\description{
Inserts text in 'Body' style at the current cursor point in Word.
}
\usage{
wdBody(text = "", paragraph = TRUE, wdapp = .R2wd)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{text}{
a test string to be written to Word.
}
  \item{paragraph}{
whether a paragraph (line return) should be added at the end.
}
  \item{wdapp}{
the handle to word (can be ignored).
}
}
\details{
Switches to 'Body' style, adds the text, and then gives a carriage
return (paragraph break) if requested.
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
\dontrun{wdBody("R2wd is a package to write MS-Word files from R")}
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{IO}
\keyword{connection}

