option explicit

sub main() ' {

    dim bytes() as byte
    bytes = strConv("Ren�", vbFromUnicode)
'   bytes = strConv("Ren�", vbUnicode)

    debug.print("Length of bytes: " & uBound(bytes) - lBound(bytes) + 1)
    dim pos as long
    for pos = lBound(bytes) to uBound(bytes)
        debug.print(chr(bytes(pos)) & " (" & bytes(pos) & ")")
    next pos

end sub ' }
