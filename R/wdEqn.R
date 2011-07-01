wdEqn<-
function (eqtext, bookmark = NULL, iknow=FALSE, wdapp = .R2wd, paragraph = TRUE){
    if (!iknow) stop("Beware: this function is dangerous (it works using sendkeys). If you wish to use it anyway you must set iknow=TRUE")
    wsshell<-comGetObject("WScript.Shell")
    if (is.null(wsshell)) wsshell<-comCreateObject("WScript.Shell")
    ## pass text to Word
    wdsel<-wdapp[["Selection"]]
    eqtext<-gsub("\\(","\\{\\(\\}",eqtext)
    eqtext<-gsub("\\)","\\{\\)\\}",eqtext)
    eqtext<-gsub("\\+","\\{\\+\\}",eqtext)
    eqtext<-gsub("\\*","\\{\\*\\}",eqtext)
    eqtext<-gsub("\\^","\\{\\^\\}",eqtext)
    eqtext<-gsub("=","\\{=\\}",eqtext)
    eqtext<-gsub("~","\\{~\\}",eqtext)
    wdapp$activate()
    wsshell$AppActivate("Microsoft Word")
    wsshell$SendKeys("%=")
    wsshell$SendKeys(eqtext)
    wsshell$SendKeys(" ~")
    wdNormal(paragraph=FALSE)
    return()
}

