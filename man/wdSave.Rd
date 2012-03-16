\name{wdSave}
\Rdversion{1.1}
\alias{wdSave}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Save the active document.
}
\description{
Save the active document. If a name is given, it is used for the new
document, otherwise Word will ask.
}
\usage{
wdSave(Name = NULL, wdapp = .R2wd)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{Name}{
    File name (if missing, Word will ask).
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
\examples{
\dontrun{
wdGet()
wdSave("This File.doc")
wdQuit()
}
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{IO}
\keyword{connection}
