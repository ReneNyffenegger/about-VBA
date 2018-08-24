option explicit

sub main()

    dim obj_1 as Foo
    dim obj_2 as Foo

    debug.print objRefCnt(obj_1) ' 0

    set obj_1 = new Foo
    debug.print objRefCnt(obj_1) ' 1

    set obj_2 = obj_1
    debug.print objRefCnt(obj_1) ' 2
    debug.print objRefCnt(obj_2) ' 2

    set obj_2 = new Foo
    debug.print objRefCnt(obj_1) ' 1
    debug.print objRefCnt(obj_2) ' 1

end sub
