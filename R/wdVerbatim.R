wdVerbatim<-
function (text = "", paragraph = TRUE, fontsize=9,fontname="Courier New",wdapp = .R2wd)
{
   wdsel <- wdapp[["Selection"]]
   wdsel$TypeParagraph()
   savestyle <- wdsel[["Style"]]
   wdsel[["Style"]] <- -67
   wdsel[["ParagraphFormat"]][["LineSpacingRule"]] <- 0
   wdsel[["Font"]][["Name"]]<-"Courier New"
   wdsel[["Font"]][["Size"]]<-fontsize
   newtext <- paste(text,collapse="\n")
   wdsel$TypeText(newtext)
   wdsel$TypeParagraph()
   wdsel[["Style"]] <- savestyle
}
