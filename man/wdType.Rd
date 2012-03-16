\name{wdType}
\Rdversion{1.1}
\alias{wdType}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
write text to word
}
\description{
Inserts text at the current cursor point in Word (by default without
messing with styles).
}
\usage{
wdType(text = "", italic=FALSE, alignment="nothing", paragraph = TRUE, wdapp = .R2wd)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{text}{
a test string to be written to Word.
}
  \item{italic}{
if \code{TRUE} then the text will be written in italics
}
  \item{alignment}{
choice of "nothing", "left", "center", "right".
}
  \item{paragraph}{
whether a paragraph (line return) should be added at the end.
}
  \item{wdapp}{
the handle to word (can be ignored).
}
}
\details{
Switches to 'Type' style, adds the text, and then gives a carriage
return (paragraph break) if requested.
}
\author{
Christian Ritter
}

%% ~Make other sections like Warning with \section{Warning }{....} ~

\examples{
\dontrun{wdType("R2wd is a package to write MS-Word files from R")}
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{IO}
\keyword{connection}

