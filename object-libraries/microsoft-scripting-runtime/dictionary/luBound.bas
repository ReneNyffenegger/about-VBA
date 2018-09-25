option explicit

sub main() ' {

    dim dict as new dictionary

    dict("k1") = "v1"
    dict("k2") = "v2"
    dict("k3") = "v3"

    debug.print "keys  - lBound: " & lBound(dict.keys ) & ", uBound: " & uBound(dict.keys )
    debug.print "items - lBound: " & lBound(dict.items) & ", uBound: " & uBound(dict.items)

end sub ' }
