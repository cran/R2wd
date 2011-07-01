wdType <-
function (text = "",
          italic = FALSE,
          alignment = "nothing",
          paragraph = TRUE,
          wdapp = .R2wd)
{
    wdsel <- wdapp[['Selection']]
    if (italic) wdsel[['Font']][['Italic']]<--1
    if (alignment=="left") wdsel[['ParagraphFormat']][['Alignment']]<-0
    if (alignment=="center") wdsel[['ParagraphFormat']][['Alignment']]<-1
    if (alignment=="right") wdsel[['ParagraphFormat']][['Alignment']]<-2
    newtext<-text
    newtext[newtext==""]<-"\n"
    newtext<-paste(newtext,collapse=" ")
    newtext<-gsub("\n ","\n",newtext)
    newtext<-gsub(" \n","\n",newtext)
    wdsel$TypeText(newtext)
    if (paragraph){
        wdsel$TypeParagraph()
    }
}

