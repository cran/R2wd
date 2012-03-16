\name{wdTitle}
\Rdversion{1.1}
\alias{wdTitle}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Add text in title style.
}
\description{
Add text in title style and add a paragraph bread if needed.
}
\usage{
wdTitle(title,
label = substring(gsub("[, .]", "_", paste("text", title, sep = "_")), 1, 16),
paragraph = TRUE, wdapp = .R2wd)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{title}{
    the text to format in title style
}
  \item{label}{
    the bookmark text (if given)
}
  \item{paragraph}{
    whether a paragraph break should be added at the end.
}
  \item{wdapp}{
    the handle to the Word Application (usually not needed).
}
}
\author{
Christian Ritter
}

%% ~Make other sections like Warning with \section{Warning }{....} ~

\examples{
\dontrun{
wdGet()
wdTitle("Title")
wdSection("Section 1",newpage=TRUE)
wdSubsection("Sub-section 1.1")
wdSubsubsection("Sub-sub-section 1.1.a")
}
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{IO}
\keyword{connection}
