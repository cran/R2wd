\name{wdStyleSetup}
\Rdversion{1.1}
\alias{wdStyleSetup}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Modify document style
}
\description{
Modify document style
}
\usage{
wdStyleSetup(style = "Normal", fontsize = 11, align=3, wdapp = .R2wd)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{style}{
    default: "Normal"
}
  \item{fontsize}{
    default 11pt
}
  \item{align}{
    type of alignment: 3 corresponds to "justified"
}
  \item{wdapp}{
    handle to the Word Application (usually not needed).
}
}
\details{
  this invokes the required methods in word to set the document style
}
\author{
Christian Ritter
}

%% ~Make other sections like Warning with \section{Warning }{....} ~

\examples{
\dontrun{
 wdGet()
wdStyleSetup()
}
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{IO}
\keyword{connection}
