option explicit

sub main() ' {

    dim D as new dictionary

    if isEmpty(D("K")) then
       debug.print "is empty"
    else
       debug.print "is not empty"
    end if

    D("K") = "val"

    if isEmpty(D("K")) then
       debug.print "is empty"
    else
       debug.print "is not empty"
    end if

end sub ' }
