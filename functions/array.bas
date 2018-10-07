option explicit

sub main() ' {

    dim ary as variant
    ary = array("zero", "one", "two", "three", "four", "five")

    dim i as long
    for i = lBound(ary) to uBound(ary)

        debug.print i & ": " & ary(i)

    next i

end sub ' }
'
' 0: zero
' 1: one
' 2: two
' 3: three
' 4: four
' 5: five
