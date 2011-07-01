wdPlot<-
function (..., plotfun = plot, method="metafile",height = 5, width = 5,pointsize = 10,
     bookmark = NULL, wdapp = .R2wd, paragraph = TRUE)
{
     wdsel <- wdapp[["Selection"]]
     wddoc <- wdapp[["ActiveDocument"]]
     if (paragraph) wdsel$TypeParagraph()
     switch(method,

	metafile={
		win.metafile(height = height, width = width, pointsize = pointsize)
     		plotfun(...)
     		dev.off()
		wdsel$Paste()

	},
	bitmap= {
		bmp(file="R2wdtemp.bmp",height=height,width=width,unit="in",res=72)
		plotfun(...)
		dev.off()
		path<-paste(getwd(),"R2wdtemp.bmp",sep="/")
		path<-gsub("/","\\\\",path)
    		wdsel[["InlineShapes"]]$AddPicture(path)
	},
	stop("method must be 'metafile' or 'bitmap'")
	)
     wdsel$MoveLeft()
     wdsel$Expand()
     if (is.null(bookmark))
         bookmark <- paste("InlineShape",
	wddoc[["InlineShapes"]][["Count"]],
             sep = "")
     wddoc[["Bookmarks"]]$Add(bookmark)
     wdsel$MoveRight()
     if (paragraph)
         wdsel$TypeParagraph()
}

