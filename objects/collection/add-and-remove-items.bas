option explicit

sub main () ' {

    dim coll as collection
    dim elem as variant

    set coll = new collection

    coll.add("one"  )
    coll.add("two"  )
    coll.add("three")
    coll.add("four" )
    coll.add("five" )
    coll.add("six"  )

    debug.print(coll(3)) ' "three"

    debug.print("going to remove element 2 and 4")

    coll.remove(2)
    coll.remove(4) ' removes element "five" because this is the 4th element after removing the 2nd element

    debug.print(coll(4)) ' "six"

end sub ' }
