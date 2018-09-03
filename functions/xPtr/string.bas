option explicit

sub main()

    dim addr as longPtr

    dim str  as string

    str = "foo bar baz"

    addr = varPtr(str)
    debug.print "varPtr(str) =         " & addr

    addr = GetMem1_(addr + 0)                        + _
           GetMem1_(addr + 1) * 256&                 + _
           GetMem1_(addr + 2) * 256& * 256&          + _
           GetMem1_(addr + 3) * 256& * 256& * 256&

    debug.print "varPtr(str) points at " & addr

    addr = strPtr(str)
    debug.print "strPtr(str) =         " & addr

end sub

' varPtr(str) =         2355884
' varPtr(str) points at 147336780
' strPtr(str) =         147336780
