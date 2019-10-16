option explicit

sub s(byVal pVal as string, byRef pRef as string) ' {

    pVal = "changed"
    pRef = "changed"

end sub ' }

sub main() ' {

    dim var1 as string
    dim var2 as string

    var1 = "some value"
    var2 = "some value"

    s var1, var2

    debug.print "var1 = " & var1 & ", var2 = " & var2

end sub ' }
'
'   Executing main prints
'
'     var1 = some value, var2 = changed
