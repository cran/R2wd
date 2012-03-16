\name{wdEqn}
\alias{wdEqn}
%- Also NEED an '\alias' for EACH other topic documented here.
\title{
Write formulas in Word
}
\description{
This function allows writing formulas/equations into Word using the
default formula editor in Word 2007 and 2010.
}
\usage{
wdEqn(eqtext, bookmark = NULL, iknow=FALSE,waitsec=2,wdapp = .R2wd, paragraph = TRUE)
}
%- maybe also 'usage' for other objects documented here.
\arguments{
  \item{eqtext}{
    a text string containing the formula
}
  \item{bookmark}{
    a text string containing the bookmark to use (otherwise default)
}
  \item{iknow}{
    acknowledgement of danger using sendkeys
}
  \item{waitsec}{
    time R waits while sendkey sends the equation
}
  \item{wdapp}{
    the handle to the Word application
}
  \item{paragraph}{
    whether to insert a paragraph after the equation
}
}
\details{
This function uses two tools to write a formula to Word. At first it
uses the COM client to write a text into the range property of the
selection and converts it to a formula. Then it uses the sendkeys method
of the wsshell to add a carriage return. This causes Word to parse the
formula and to substitute LaTeX like tokens such as \code{\\alpha} by their
corresponding formula equivalents (the greek letter alpha). In
principle, it should be possible to achieve the parsing of the formula
by invoking the BuildUp method of the omaths(1) object, but this does
not work in the current versions of Word. The route via Sendkeys is
fragile, so handle with caution.
}
\value{
  none
}

%% ~Make other sections like Warning with \section{Warning }{....} ~

\examples{
\dontrun{
wdGet()
wdTitle("Try a formula")
## note that double \\ has to be used instead of \
## note that the , the following lines
## must be pasted into the console and run there.
## running them from an Emacs/ESS buffer won't work
wdEqn("\\\\alpha \\\\over ( \\\\beta +3 ) ")
wdEqn("\\\\alpha \\\\over (( \\\\beta +3 ))")
wdEqn("\\\\matrix(\\\\alpha &0@0&1 ) ")
## note: the handling of parentheses "(" is a bit peculiar in Word
## also experiment with the handling of empty spaces
}
}
% Add one or more standard keywords, see file 'KEYWORDS' in the
% R documentation directory.
\keyword{ IO }
