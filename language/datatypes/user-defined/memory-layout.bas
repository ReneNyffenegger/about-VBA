option explicit

type T ' {
     t_1 as long
     t_2 as long
     t_3 as long
end  type ' }

type U ' {
     t_     as T
     u_1    as long
     u_2    as long
end type ' }

sub main() ' {

    dim tt as T

  ' debug.print varPtr(tt)

  '
  ' The following line prints 0. Thus, the address of
  ' first element in a type is the same as the
  ' address of the type-instance itself:
  '
    debug.print varPtr(tt.t_1) - varPtr(tt    )

  '
  ' The following two lines each print 4.
  ' This indicates that no memory is wasted between the
  ' elements of a type (a long is four bytes).
  '
    debug.print varPtr(tt.t_2) - varPtr(tt.t_1)
    debug.print varPtr(tt.t_3) - varPtr(tt.t_2)

    dim uu as U

  ' debug.print varPtr(uu)

  '
  ' The following statement prints 0.
  ' Again, the address of the type-instance is the
  ' same as the address of its first element, even
  ' if this element happens to be an instance of a type.
  '
    debug.print varPtr(uu.t_ ) - varPtr(uu    )

  '
  ' The following lines prints 12. Thus, the first
  ' element occupies 12 bytes which corresponds to the
  ' 3 longs in T (3 * 4 = 12)
  '
    debug.print varPtr(uu.u_1) - varPtr(uu    )

  '
  ' As now expected, the following statement prints 4:
  '
    debug.print varPtr(uu.u_2) - varPtr(uu.u_1)

end sub ' }
