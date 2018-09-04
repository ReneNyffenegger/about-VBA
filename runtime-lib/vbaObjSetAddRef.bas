sub main()

    dim coll as new collection
    dim obj  as object

    coll.add "foo"
    coll.add "bar"

  '
  ' The following call seems to be equivalent to
  '
  '      set obj = coll
  '
    call vbaObjSetAddRef(obj, objPtr(coll))

  '
  ' obj now refers to the same object as coll. Thus,
  ' adding "baz" to obj is ...
  '
    obj.add "baz"

  '
  ' ... reflected in coll. Iterating over coll's
  ' items prints baz as well.
  '
    dim i as variant
    for each i in coll
        debug.print i
    next i

end sub
