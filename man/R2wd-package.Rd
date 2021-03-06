\name{R2wd-package}
\Rdversion{1.1}
\alias{R2wd-package}
\alias{R2wd}
\docType{package}
\title{
Write MS-Word documents from R.
}
\description{
R2wd uses a COM client to communicate with MS-Word via the COM interface
(Windows only). Two COM clients are supported: "rcom" (Statconn DCOM via
Rscproxy) and RDCOMClient (from the omegahat distribution). R2wd can be
seen as a  collection of convenience wrappers to calls
of the COM client.}
\details{
\tabular{ll}{
Package: \tab R2wd\cr
Type: \tab Package\cr
Version: \tab 1.5.0\cr
Date: \tab 2012-02-28\cr
License: \tab GPL-2\cr
LazyLoad: \tab yes\cr
}
If Word is not already running, wdGet() opens a new Word document,
otherwise, it establishes a COM handle to the instance which is already
running. If the parameter "filename" is specified, wdGet will check
whether this file is already open or (if not) try to open it.
The functions wdTitle, wdHeader, wdType, wdBody, and wdParagraph can be
used to inject text elements into Word. Moreover, bookmarks can be added
via wdInsertBookmarks and wdGoToBookmark allows to navigate among the
bookmarks which also exist. There is another set of convenience
functions, wdSection, wdSubsection, and wdSubsubsection which insert
headers of level 1, 2, or 3, start new 'Sections' in Word, and add
bookmarks.

Graphs and dataframes can be inserted into Word, by the wdPlot, wdTable
commands. The wdTable command takes a dataframe or an array as
arguments, creates a Word table of the appropriate dimensions and
injects the content of the dataframe or array into it. It then formats
the table in Word using elementary formating elements.

There are a few convenience functions for carrying out various typical
functions in Word, such as undo, page setup, verbatim, etc.

The functions wdApplyTheme and wdApplyTemplate allow to work with themes
and templates.
}
\author{
Christian Ritter

Maintainer: Christian Ritter <R2wd@ridaco.be>
}
\keyword{ package }
\examples{
\dontrun{
## rcom needs to be installed
if (!require(rcom)) warning("Install rcom first")
## This initializes the hook to Word
## if Word is already running it connects to it via StatconnDcom
## if not, it uses StatconnDcom to open word to create a new document
## and to connect to it.
wdGet()
## Then there is a set of small functions which implement elementary
## functionality:
wdTitle("R2wd, A Package for writing Word Documents from R",label="R2wd")
##
wdSection("Introduction")
##
wdBody("This is an example on how to use the R2wd package.")
wdWrite("wdWrite continues writing in the same style. If you set paragraph to FALSE, it doesn't add a paragraph return at the end. For example, you can use to add a footnote",paragraph=FALSE)
wdInsertFootnote("Footnotes have two arguments, the footnote text and the footnote reference. By default, word creates the reference automatically, but you can also choose what you want.","whatUwant")
wdBody(". As it is shown here")
##
## We can also plot data. The following line uses the default plot command
##
wdPlot(1:100,sin((1:100)/10),type="l",main="a test graph")
##
## The plot command can be exchanged for something else
##
wdPlot(mtcars,plotfun=pairs,main="a test graph",height=6,width=6,pointsize=8)
##
## we can also do something more elaborate (requires latticeExtra)
##
if (!require(latticeExtra)) warning("Install latticeExtra first")
fun<-function(...) print(useOuterStrips(xyplot(...)))
wdPlot(mpg~hp|cyl+carb,data=mtcars,plotfun=fun,height=6,width=6)
##
## There is also rudimentary way to transfer data frames.
##
wdTable(mtcars)
##
## It may be useful to format them first
##
wdTable(format(mtcars))
##
## Finally, we can apply themes and templates
##
wdApplyTheme("Quad 100")
wdApplyTheme("BlueCalm 100")
wdApplyTheme("None")
wdSave("MyDoc")
##
wdQuit()
##
## Another example on how to wrap the definition of a plot into a function and call it from wdPlot
##
xseries = ts(1:100) #make the series of the wind stress
yseries = ts(1:100)     #time series objects
myplot<-function(x,y){
   	par(mfrow=c(3,3), mar=c(2.5,4,4,1),oma=c(2,0,0,0))   #set up plot area
	for (h in 0:8) {     #loop through lags 0-8
      		plot.ts(xy.lines=F,xy.labels=F,lag(x,-h),y,main=paste("x(t-",h,")",sep=""),
		ylab=expression(x[t]),xlab="")
	mtext("Yearly",SOUTH<-1, line=0.0,adj=0, cex=1, col="Black", outer=TRUE)
 }
}
wdPlot(x,y,plotfun=myplot,height=6,width=6,pointsize=9,bookmark="myplot",paragraph=TRUE)
##
}
}
