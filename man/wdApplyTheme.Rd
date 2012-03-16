\name{wdApplyTheme}
\Rdversion{1.1}
\alias{wdApplyTheme}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Apply a word template.
}
\description{
Apply a word template.
}
\usage{
wdApplyTheme(theme, wdapp = .R2wd)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{theme}{
The name of the theme.
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
\author{
Christian Ritter
}
\note{
This function is just a sketch on what could be done.
}


\examples{

\dontrun{
## apply a default theme with a blue background.
wdApplyTheme("BlueCalm 100")
## remove the current theme
wdApplyTheme("None")
}
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{IO}
\keyword{connection}
