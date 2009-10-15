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
wdSave("This File.doc")
wdQuit()
}
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{IO}
\keyword{connection}
