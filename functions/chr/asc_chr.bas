option explicit

sub main() ' {

    dim charNo as integer

    for charNo = asc("a") to asc("z")
        debug.print "chr(" & charNo & ") = " & chr(charNo)
    next charNo

end sub ' }
