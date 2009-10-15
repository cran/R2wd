\name{wdGet}
\Rdversion{1.1}
\alias{wdGet}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Gets a handle to Word.
}
\description{
wdGet checks first if Word is already running. If this is not the case,
it will start Word, add a document and then return the handle to the
Word application. If Word is already running and has an active document,
it will just return the handle.
}
\usage{
wdGet(visible = TRUE)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{visible}{
whether the Word application should be visible.
}
}
\details{
%%  ~~ If necessary, more details than the description above ~~
}
\value{
a handle to the COM object refering to the Word application.
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
wdTitle("R2wd, a package to write Word documents from R")
}
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{IO}
\keyword{connection}

