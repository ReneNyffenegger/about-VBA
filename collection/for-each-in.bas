sub main()

    dim c as new collection

    c.add "foo"
    c.add "bar"
    c.add "baz"

    for each i in c
        debug.print i
    next

end sub
