option explicit

sub main() ' {

    dim dict as new dictionary

    dim ob as obj

    set ob = new obj : ob.init 42, "Hello world" : dict.add "one"  , ob
    set ob = new obj : ob.init 99, "ninety-nine" : dict.add "two"  , ob
    set ob = new obj : ob.init 21, "1+2+3+4+5+6" : dict.add "three", ob

    iterateOverKeys dict

end sub ' }

sub iterateOverKeys(d as dictionary) ' {

    dim k as variant
    for each k in d ' {
        debug.print k & " -> " & d(k).num & ": " & d(k).txt
    next k ' }

end sub ' }
