wdBody <-
function (text = "", paragraph = TRUE, wdapp = .R2wd)
{
    wdsel <- wdapp[['Selection']]
    savestyle<-wdsel[["Style"]]
    wdsel[["Style"]]<--67
    wdsel$TypeText(text)
    if (paragraph){
        wdsel$TypeParagraph()
        wdsel[["Style"]]<-savestyle}
}

