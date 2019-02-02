option explicit

declare function rtcErrObj lib "VBE7.dll" alias "rtcErrObj"() as object

sub main() ' {

    dim flOldProtect     as long
    dim vbe7_dll         as long

    dim objErr           as object

    set objErr = rtcErrObj()

    debug.print "objPtr(objErr) = " & objPtr(objErr)
    debug.print "objPtr(err   ) = " & objPtr(err   )

end sub ' }
