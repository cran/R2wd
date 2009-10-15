wdHeading <-
function (level=1,text = "", paragraph = TRUE, wdapp = .R2wd)
{
    wdsel<-wdapp[['Selection']]
    wdsel[["Style"]]<--1-level
    wdsel$TypeText(text)
    if (paragraph){
        wdsel$TypeParagraph()
        wdsel[["Style"]]<- -1
    }
    invisible()
}

