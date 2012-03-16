\name{wdSetFont}
\alias{wdSetFont}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Set font in Word.
}
\description{
This function allows to set the font type and the font size of the
active Word document.
}
\usage{
wdSetFont(fontname = NULL, fontsize = NULL, bold=NULL, italic=NULL,wdapp = .R2wd)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{fontname}{
A font name Word knows about, such as Arial, Times New Roman, etc.
}
  \item{fontsize}{
The point size of the font.
}
  \item{bold}{
The point size of the font.
}
  \item{italic}{
The point size of the font.
}
  \item{wdapp}{
The handle to Word.
}
}
\details{
This function uses the handle .R2wd to talk to Word and to change font
name and size at the current position of the cursor in the active document.
}
\author{
Christian Ritter
}
%% ~Make other sections like Warning with \section{Warning }{....} ~

\examples{
\dontrun{
wdGet()
wdTitle("Title")
wdWrite("This is in normal font")
wdSetFont(fontname="Garamond",fontsize=14,bold=TRUE,italic=TRUE)
wdWrite("This is in 14pt Garamond")
wdNormal("This switches back to Normal")
}
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{IO}
\keyword{connection}
