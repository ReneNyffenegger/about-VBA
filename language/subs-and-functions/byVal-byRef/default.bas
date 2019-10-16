option explicit

sub s(p as string) ' {
    debug.print "p = " & p
    p = "changed"
end sub ' }

sub main() ' {

    dim n as string

    n = "with"     ' With parantheses, the value is passed «byVal»
    s(n)           ' The assignment of a value to the parameter within the sub
    s(n)           ' does not affect the value of n

    n = "without"  ' But… without the parantheses, the value of
    s n            ' n is passed «byRef», that is, it can be
    s n            ' changed within the sub

    n = "call"     ' If the call statement is used, the value of
    call s(n)      ' the parameter is also passed «byRef» (although
    call s(n)      ' parantheses are required).

end sub ' }

'
'   Executing main prints
'
'     p = with
'     p = with
'     p = without
'     p = changed
'     p = call
'     p = changed
