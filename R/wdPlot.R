wdPlot <-
function(...,plotfun=plot,height=5,width=5,pointsize=10,bookmark=NULL,wdapp=.R2wd,paragraph=TRUE){
    win.metafile(height=height,width=width,pointsize=pointsize)
    plotfun(...)
    dev.off()
    wdsel<-wdapp[['Selection']]
    wddoc<-wdapp[['ActiveDocument']]
    wdsel$TypeParagraph()
    wdsel$Paste()
    wdsel$MoveLeft()
    wdsel$Expand()
    if (is.null(bookmark))
        bookmark<-paste("InlineShape",wddoc[["InlineShapes"]][["Count"]],sep="")
    wddoc[["Bookmarks"]]$Add(bookmark)
    wdsel$MoveRight()
    if(paragraph) wdsel$TypeParagraph()
}

