\name{wdSectionBreak}
\Rdversion{1.1}
\alias{wdSectionBreak}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Insert a section break.
}
\description{
Insert a section break. Depending on the setting of \code{continuous}
start a new page for this.
}
\usage{
wdSectionBreak( continuous = TRUE, bookmark = NULL,wdapp = .R2wd)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{continuous}{
    if FALSE, the Word section will start on a new page
}
  \item{bookmark}{
    bookmark text
}
  \item{wdapp}{
    handle to Word application (usually not needed)
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
%% none
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{IO}
\keyword{connection}

