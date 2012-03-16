\name{wdGet}
\Rdversion{1.1}
\alias{wdGet}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Gets a handle to Word (using RDCOMClient).
}
\description{
wdGet checks first if Word is already running. If this is not the case,
it will start Word, add a document and then return the handle to the
Word application. If Word is already running and has an active document,
it will just return the handle.
}
\usage{
wdGet(filename = NULL, path = "", method="rcom",visible = TRUE)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{filename}{
the filename of an existing word document (if null, make a new one)
}
  \item{path}{
the path to an existing word document
}
  \item{method}{
the COM client. Either "rcom" (default) or "RDCOMClient"
}
  \item{visible}{
whether the Word application should be visible.
}
}
\value{
a handle to the COM object refering to the Word application.
}
\author{
Christian Ritter
}

%% ~Make other sections like Warning with \section{Warning }{....} ~

\examples{
\dontrun{
wdGet()
wdTitle("R2wd, a package to write Word documents from R")
}
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{IO}
\keyword{connection}

