\name{wdNewDoc}
\Rdversion{1.1}
\alias{wdNewDoc}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Start a new document (when Word is running).
}
\description{
This adds a new document to a running Word Application. If a name is
given, the new document is saved under the given name.
}
\usage{
wdNewDoc(name = NULL, wdapp = .R2wd)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{name}{
    the optional name of the file.
}
  \item{wdapp}{
    the handle to word (usually not needed)
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
\code{\link{wdGet}}
}
\examples{
\dontrun{
wdGet()
wdNewDoc("This.doc")
wdQuit()
}
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{IO}
\keyword{connection}
