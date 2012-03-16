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
wdTable(data, caption = "", caption.pos="below",bookmark = NULL, pointsize = 9, padding = 5, autoformat = 1, row.names=TRUE,align = if (row.names) c("l",rep("r", ncol(data))) else c(rep("r", ncol(data))),hlines=NULL,wdapp = .R2wd)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{data}{
    a data frame or an array
}
  \item{caption}{
    the caption to use in Word
}
  \item{caption.pos}{
    the caption position "below" or "above"
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
  \item{row.names}{
    whether row names should be printed (default=TRUE)
}
  \item{align}{
    alignment instruction. Default: c("l",rep("r",ncol(data))): align
    first column left and the others right. "c" stands for
    centering. The character "|" can be inserted where the user wishes a
    vertical line to be drawn.
}
  \item{hlines}{
    a character vector of length nrow(data)+1 containing "b","t","bt",or
    "n" to indicate whether a horizontal line should be drawn below, on
    top, or both. "n" means none.
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
\author{
Christian Ritter
}

%% ~Make other sections like Warning with \section{Warning }{....} ~

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
