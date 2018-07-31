option explicit

sub main()

    dim c as new collection

    c.add "foo"
    c.add "bar"
    c.add "baz"

    dim i as variant ' With option explicit, i must be declared.

    for each i in c
        debug.print i
    next

end sub
