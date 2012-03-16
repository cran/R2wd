\name{wdVerbatim}
\alias{wdVerbatim}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
write verbatim text to word (for example model output).
}
\description{
write single spaced text in (by default) Courier font to word. This is
useful to render R output 'as is'.
}
\usage{
wdVerbatim(text = "", paragraph = TRUE, fontsize = 9, fontname = "Courier New", wdapp = .R2wd)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{text}{
the text to write
}
  \item{paragraph}{
whether the chunk should be closed by a paragraph return.
}
  \item{fontsize}{
size of the font
}
  \item{fontname}{
name of the font
}
  \item{wdapp}{
handle to the Word application (usually not used).
}
}
\details{
Saves the current style. Wwitches to Courier New, single spaced, 9 pt
(by default) and renders the text. Switches back to the previous style.
}
\author{
Christian Ritter
}
\note{
This is almost the similar to \code{wdBody} and \code{wdWrite}.
}

%% ~Make other sections like Warning with \section{Warning }{....} ~

\examples{
\dontrun{
wdGet()
tt<-capture.output(summary(lm(mpg~hp,data=mtcars)))
wdVerbatim(tt)
  }
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{IO}
\keyword{connection}
