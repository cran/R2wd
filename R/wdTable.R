wdTable <-
function(data,caption="",bookmark=NULL,pointsize=9,padding=5,wdapp=.R2wd){
  wdtab<-wdapp[['ActiveDocument']][['Tables']]
  wdsel<-wdapp[['Selection']]
  wddoc<-wdapp[['ActiveDocument']]
  ## set endbookmark after the table
  wdsel$TypeParagraph()
  wdInsertBookmark("R2wdEndmark")
  bookmarkcounter<-wddoc[['Bookmarks']][['Count']]
  wdsel$MoveUp()
  wdselrg<-wdsel[['Range']]
  nr<-nrow(data)
  nc<-ncol(data)
  tab<-wdtab$Add(wdselrg,nr+1,nc+1)
  for (j in (1:nc)){
    cellrg<-comInvoke(tab,"Cell",1,j+1)[['Range']]
    cellrg[['Text']]<-colnames(data)[j]
  }
  for (i in (1:nr)){
    cellrg<-comInvoke(tab,"Cell",i+1,1)[['Range']]
    cellrg[['Text']]<-row.names(data)[i]
  }
  for (i in 1:nr)
    for (j in 1:nc){
    cellrg<-comInvoke(tab,"Cell",1+i,j+1)[['Range']]
    cellrg[['Text']]<-data[i,j]
  }
  tab[['AutoFitBehavior']]<-"wdAutoFitContent"
  tab$AutoFormat()
  tab[['Rows']][['Height']]<-pointsize+padding
  tab[['Rows']][['HeightRule']]<-2
  tab[['Range']][['ParagraphFormat']][['Alignment']]<-2
  tab[['Range']][['Cells']][['VerticalAlignment']]<-1
 if(is.null(bookmark)) bookmark<-paste("bookmark",bookmarkcounter+1,sep="")
  wdInsertBookmark(bookmark)
  wddoc[['Bookmarks']]$Item(bookmarkcounter)$Select()
  #wdapp[['Selection']]$InsertCaption("Table",caption)
  wdParagraph()
  return()
}

