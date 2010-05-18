wdGet <-
function (filename=NULL,path="",doc=NULL,visible=TRUE)
{
    require(rcom)
    comRegisterServer()
    tmp<-comGetObject("Word.Application")
    if (is.null(tmp)) tmp<-comCreateObject("Word.Application")
    if (is.null(filename)){
        if (is.null(tmp[['ActiveDocument']])) tmp[['Documents']]$Add()} else
        if(!(tmp[['ActiveDocument']][['Name']]==filename)) tmp$Open(paste("path",filename,sep="/"))
    if (visible) tmp[['visible']]<-TRUE
    .R2wd<<-tmp
    invisible(tmp)
}

