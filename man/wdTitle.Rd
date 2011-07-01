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
%%  ~~who you are~~
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
