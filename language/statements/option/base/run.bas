option explicit

sub main() ' {

    dim array_0() as variant
    dim array_1() as variant

    array_0 = arrayWithoutOptionBase
    array_1 = arrayWithOptionBase

    debug.print "array_0 starts with index " & lBound(array_0) & " and ends with index " & uBound(array_0)
    debug.print "array_1 starts with index " & lBound(array_1) & " and ends with index " & uBound(array_1)
    debug.print ""
    debug.print "array_0(2) = " & array_0(2)
    debug.print "array_1(2) = " & array_1(2)

end sub ' }
'
' Program writes
'
'   array_0 starts with index 0 and ends with index 3
'   array_1 starts with index 1 and ends with index 4
'
'   array_0(2) = 2
'   array_1(2) = two
