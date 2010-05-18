\name{wdSection}
\Rdversion{1.1}
\alias{wdSection}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Start a new section of the document.
}
\description{
  Start a section by:
  Adding a section break,
  switching to Heading1 style,
  adding the section title,
  adding a bookmark,
  adding a paragraph break,
  switching to Normal style.
 }
\usage{
wdSection(title, label = gsub(" ", "_", paste("sec", title, sep = "_")),
newpage = FALSE, wdapp = .R2wd)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{title}{
    section title
}
  \item{label}{
    bookmark text (will be generated automatically if missing)
}
  \item{newpage}{
    whether the section should start on a new page (doesn't work)
}
  \item{wdapp}{
    handle to Word Application (usually not needed).
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
\dontrun{
wdGet()
wdTitle("R2wd a package to ... ")
wdSection("Introduction",newpage=TRUE)

}
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{IO}
\keyword{connection}
