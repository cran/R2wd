wdNormal <-
function (text = "", paragraph = TRUE, wdapp = .R2wd)
{
    wdsel <- wdapp[['Selection']]
    wdsel[['Style']]<- -1
    wdsel$TypeText(text)
    if (paragraph) wdsel$TypeParagraph()
}

