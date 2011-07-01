wdEqn<-
function (eqtext, bookmark = NULL, wdapp = .R2wd, paragraph = TRUE){
    wsshell<-comGetObject("WScript.Shell")
    if (is.null(wsshell)) wsshell<-comCreateObject("WScript.Shell")
    ## pass text to Word
    wdsel<-wdapp[["Selection"]]
    objrange<-wdsel[["Range"]]
    objrange[["Text"]]<-eqtext
    ## turn into equation
    wdsel[["OMaths"]]$Add(objrange)
    wdsel[["OMaths"]]$Item(1)$BuildUp()
    wdsel$EndKey()
    wdapp$activate()
    wsshell$AppActivate("Microsoft Word")
    wsshell$SendKeys(" ~")
    return()
}

