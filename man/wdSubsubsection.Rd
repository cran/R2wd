\name{wdSubsubsection}
\Rdversion{1.1}
\alias{wdSubsubsection}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Start a new Word section, add a heading and a bookmark.
}
\description{
Start a subsection (see \code{\link{wdSection}} for details.
}
\usage{
wdSubsubsection(title, label = gsub(" ", "_",
paste("subsec", title, sep = "_")),
newpage = FALSE, wdapp = .R2wd)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{title}{
the sub-sub-section title.
}
  \item{label}{
the bookmark.
}
  \item{newpage}{
whether the section should start on a new page.
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
