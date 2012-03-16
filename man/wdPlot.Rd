\name{wdPlot}
\Rdversion{1.1}
\alias{wdPlot}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
create an R plot and paste it into word.
}
\description{
By default, this uses the \code{plot} function to create the plot
according to the arguments given as \code{...}. The \code{plot} function
can be replaced by another function which creates a graph.
%%  ~~ A concise (1-5 lines) description of what the function does. ~~
}
\usage{
wdPlot(..., plotfun = plot, caption="",method="metafile",height = 5, width = 5,
pointsize = 10, bookmark = NULL, wdapp = .R2wd, paragraph = TRUE)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{\dots}{
    the arguments to the plot function
}
  \item{plotfun}{
    the plot function (by default \code{plot})
}
  \item{caption}{
    figure caption
}
  \item{method}{
    the graphics device type (\code{metafile} or \code{bitmap})
}
  \item{height}{
    the height of the plot in R units (commonly inches)
}
  \item{width}{
    the width of the plot in R units (commonly inches)
}
  \item{pointsize}{
    the pointsize of the plot
}
  \item{bookmark}{
    the bookmark text (if missing, a default will be created)
}
  \item{wdapp}{
    the handle to the Word Application (usually not needed)
}
  \item{paragraph}{
    whether a paragraph is given after inserting the plot.
}
}
\details{
The plot is transfered as a windows metafile.
}
\author{
Christian Ritter
}

%% ~Make other sections like Warning with \section{Warning }{....} ~

\examples{
\dontrun{
wdGet()
## a basic call using the default metafile device
wdTitle("R2wd: plotting")
wdPlot(1:100,sin(1:100),type="l",bty="l")
## a call using a ggplot function which uses semi-transparency
## this requires the bitmap device (the metafile device is nicer but doesn't render semi-transparency)
require(ggplot2)
funny<-function(){
c <- ggplot(mtcars, aes(qsec, wt))
print(c + stat_smooth())
}
wdPlot(plotfun=funny,method="bitmap")
}
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{IO}
\keyword{connection}
