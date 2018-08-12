option explicit

sub main()

    dim obj_1 as tq84
    dim obj_2 as tq84

    debug.print "set obj_1 = new tq84"
    set obj_1 = new tq84

  '
  ' Setting obj_1's member m to foo:
  '
    obj_1.m = "foo"

  '
  ' Setting obj_2 to obj_1 does not create a new
  ' object (the operator new is not involved)
  ' obj_2 is just an alias to obj_1.
  '
    set obj_2 = obj_1

  '
  ' Changing obj_1's member m to bar.
  '
    obj_1.m = "bar"


  '
  ' Because obj_2 is an alias to obj_1,
  ' obj_2.m is equal to obj_1.m. The
  ' following line prints
  '    obj_2.m = bar
  '
    debug.print "obj_2.m = " & obj_2.m

  '
  ' Setting obj_1 to nothing does not yet
  ' invoke the destructor class_terminate.
  ' This is because obj_2 still points
  ' to the same object
  '
    set obj_1 = nothing

  '
  ' However, setting the other reference
  ' to the object to nothing will cause the
  ' class's destructor to be called.
  '
    set obj_2 = nothing

    debug.print "exiting main"

end sub

sub a(ad as long)
    debug.print "ad = " & ad
end sub
