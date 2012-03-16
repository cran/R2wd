\name{wdUndo}
\Rdversion{1.1}
\alias{wdUndo}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Save the active document.
}
\description{
Save the active document. If a name is given, it is used for the new
document, otherwise Word will ask.
}
\usage{
wdUndo(times = 1, wdapp = .R2wd)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{times}{
    the \code{times} argument of the Undo method of the word document
}
  \item{wdapp}{
    The handle to the Word Application (usually not needed).
}
}
\details{
Saves the active document to the name if given or Word will ask for a name.
}
\author{
Christian Ritter
}

%% ~Make other sections like Warning with \section{Warning }{....} ~

\examples{
\dontrun{
wdGet()
wdTitle("By default, writing a title represents 5 calls to word: Style, type text,
inserting a bookmark, typing a paragraph, and setting the style back to normal")
# we can undo it by invoking the Undo method on the document five times
wdUndo(5)
wdQuit()
}
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{IO}
\keyword{connection}
