\name{wdApplyTemplate}
\Rdversion{1.1}
\alias{wdApplyTemplate}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Apply a word template.
}
\description{
Apply a Word template.
}
\usage{
wdApplyTemplate(filename, wdapp = .R2wd)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{filename}{
The name of the template.
}
  \item{wdapp}{
The handle to the Word application (see details).
}
}
\details{
This function applies the given Word template. Indicating the handle to
the Word application is usually not necessary (will be tacitely
initialized by the \code{wdGet} function and assigned to the session
variable .R2wd.
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
This function is just a sketch on what could be done. Someone should
write a few interesting templates to mirror typical LaTEX ones, such as
article, report, etc.
}

%% ~Make other sections like Warning with \section{Warning }{....} ~

\seealso{
%% ~~objects to See Also as \code{\link{help}}, ~~~
}
\examples{

}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.

\keyword{IO}
\keyword{connection}
