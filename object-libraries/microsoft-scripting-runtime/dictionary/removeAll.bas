option explicit

sub main() ' {

    dim dict as new dictionary

    dict("k1") = "v1"
    dict("k2") = "v2"
    dict("k3") = "v3"

    debug.print "elements in dict: " & dict.count

    dict.removeAll

    debug.print "elements in dict: " & dict.count
end sub ' }
