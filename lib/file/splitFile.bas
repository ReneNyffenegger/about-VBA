option explicit

sub splitFile(fileName as string, chunkSize as long, optional overwrite as boolean = false)

    dim buf()    as byte
    dim fnrIn    as integer
    dim fnrOut   as integer
    dim fileSize as long
    dim pos      as long
    dim splitNr  as long

    on error goTo err_

    fnrIn = freeFile

    open fileName for binary access read as #fnrIn
    fileSize = LOF(fnrIn)

    reDim buf(1 to chunkSize)

    splitNr = 0
    for pos = 1 to fileSize step chunkSize
        splitNr = splitNr + 1

        if splitNr*chunkSize > fileSize then
           reDim buf(1 to fileSize - (splitNr-1)*chunkSize)
        end if

        get #fnrIn, pos, buf

        fnrOut = FreeFile

        if overwrite then
            if len(Dir$(fileName & "." & splitNr)) > 0 Then
                kill    fileName & "." & splitNr
            end if
        end if

        open cStr(fileName & "." & splitNr) for binary access write as #fnrOut
        put   #fnrOut, , buf
        close #fnrOut

    next

    close #fnrIn
    exit sub

err_:
    msgBox err.description
    close #fnrIn
    close #fnrOut
end sub
