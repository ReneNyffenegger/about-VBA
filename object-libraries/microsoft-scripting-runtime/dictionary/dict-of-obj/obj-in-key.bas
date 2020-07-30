option explicit

sub main() ' {

    dim dict as new dictionary

    dim ob as obj

    set ob = new obj : ob.init 42, "Hello world" : dict.add ob, "foo"
    set ob = new obj : ob.init 99, "ninety-nine" : dict.add ob, "bar"
    set ob = new obj : ob.init 21, "1+2+3+4+5+6" : dict.add ob, "baz"

    iterateOverKeys dict

end sub ' }

sub iterateOverKeys(d as dictionary) ' {

    dim k as variant
    for each k in d ' {
        debug.print k.num & ": " & k.txt & " -> " & d(k)
    next k ' }

end sub ' }
