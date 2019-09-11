option explicit

sub main() ' {

    #if vba6 then
        debug.print "vba6 is defined"
    #else
        debug.print "vba6 is not defined"
    #end if

    #if vba7 then
        debug.print "vba7 is defined"
    #else
        debug.print "vba7 is not defined"
    #end if

end sub ' }
