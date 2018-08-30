option explicit

sub s(p as string)
    select case p
           case "1": debug.print "one"
           case "2": debug.print "two"
           case "3": debug.print "three"
    end select
    p = "3"
end sub

sub main()

    dim n as string

    n = "1"  ' With paranthesis, the value is passed "byVal":
    s(n)     ' The change of the parameter within the sub
    s(n)     ' does not affect the value of n

    n = "2"  ' But... without the paranthesis, the value of
    s n      ' n is passed "byRef", that is, the value of
    s n      ' n is changed within the sub.

end sub
