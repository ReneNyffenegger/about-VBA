option explicit

sub s(byRef p as string) ' {
 '
 '  Note: the sub receives the paramter p «byRef» and
 '  does not change it.
 '
    debug.print "s was called with p = " & p

end sub ' }

sub main() ' {

    call s(42)

    dim var as integer
    var = 42

    call s(var)

end sub ' }
