wdGet <-
function (visible=TRUE) 
{
    require(rcom)
    comRegisterServer()
    tmp<-comGetObject("Word.Application")
    if (is.null(tmp)) tmp<-comCreateObject("Word.Application")
    if (is.null(tmp[['ActiveDocument']])) tmp[['Documents']]$Add()
    if (visible) tmp[['visible']]<-TRUE
    .R2wd<<-tmp
    invisible(tmp)
}

