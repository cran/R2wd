\name{wdTable}
\Rdversion{1.1}
\alias{wdTable}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Write a dataframe or an array as a Word table.
}
\description{
Make a Word table to the adequate dimensions, fill it, format it, and
add a bookmark (and a caption - doesn't work yet).
%%  ~~ A concise (1-5 lines) description of what the function does. ~~
}
\usage{
wdTable(data, caption = "", bookmark = NULL, pointsize = 9, padding = 5, autoformat = 1, align = c("l", rep("r", ncol(data))), wdapp = .R2wd)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{data}{
    a data frame or an array
}
  \item{caption}{
    the caption to use in Word (doesn't work yet)
}
  \item{bookmark}{
    the bookmark to use in Word
}
  \item{pointsize}{
    the pointsize in Word
}
  \item{padding}{
    how much cell-padding (in points)
}
  \item{autoformat}{
    which of the Word autoformats to use (try 1, 2, 3)
}
  \item{align}{
    alignment instruction. Default: c("l",rep("r",ncol(data))): align
    first column left and the others right. "c" stands for centering.
}
  \item{wdapp}{
    the handle to the Word Application (usually not needed).
}
}
\details{
Creates a Word table to the adequate dimensions and inserts it between
two paragraph breaks. Fills it with dataframe or array (using the clipboard).
Creates a bookmark and adds a caption. Uses word autoformating and
aligns columns according to user input.
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
wdTitle("The example mtcars")
wdTable(format(mtcars))
}
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{IO}
\keyword{connection}
