option explicit

sub main()

    dim ex_1 as Iexample
    dim ex_2 as Iexample

    set ex_1 = new Foo
    set ex_2 = new Bar

    ex_1.aSub
    ex_2.aSub

    debug.print ex_1.aFunc(11)
    debug.print ex_2.aFunc(11)

end sub
