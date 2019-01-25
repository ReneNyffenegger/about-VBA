option explicit

type T ' {
     t_lng_1  as long
     t_byt_1  as byte
     t_byt_2  as byte
     t_lng_2  as long
     t_byt_3  as byte
     t_int_1  as integer
     t_obj_1  as object
     t_obj_2  as object
end  type ' }

type U ' {
     u_t___1  as T
     u_lng_1  as long
end type ' }

sub main() ' {

    dim tt as T

  '
  ' The following line prints 0. Thus, the address of
  ' first element in a type is the same as the
  ' address of the type-instance itself:
  '
    debug.print "t_lng_1: " & varPtr(tt.t_lng_1) - varPtr(tt        )     ' 0

  '
  ' The next element in T, t_byt_1, is four bytes apart
  ' from the first element. This is because t_lng_1
  ' occupies 4 bytes.
  '
    debug.print "t_byt_1: " & varPtr(tt.t_byt_1) - varPtr(tt.t_lng_1)     ' 4

  '
  ' The next element (t_byt_2) is also a byte. It is only
  ' one byte apart from t_byt_1, because t_byt_1 occupies
  ' one byte only.
  '
    debug.print "t_byt_2: " & varPtr(tt.t_byt_2) - varPtr(tt.t_byt_1)     ' 1

  '
  ' Interestingly, the next element (t_lng_2) is three
  ' bytes apart.
  ' This is because VBA tries to align members of type long
  ' on 4 byte boundaries:
  '
    debug.print "t_lng_2: " & varPtr(tt.t_lng_2) - varPtr(tt.t_byt_2)     ' 3

  '
  ' t_lng_2 occupies 4 bytes, hence t_byt_3 is four
  ' bytes apart:
  '
    debug.print "t_byt_3: " & varPtr(tt.t_byt_3) - varPtr(tt.t_lng_2)    '  4

  '
  ' VBA also tries to align integers on 4 byte boundaries.
  ' Hence, t_int_1 is four bytes apart from t_byt_3:w

  '
    debug.print "t_int_1: " & varPtr(tt.t_int_1) - varPtr(tt.t_byt_3)    '  2


  '
  ' The next member, an object, is aligned on a four
  ' byte boundary as well. Since t_int_1 occupies two
  ' bytes only, t_obj_1 is two bytes apart from t_int_1:
  '
    debug.print "t_obj_1: " & varPtr(tt.t_obj_1) - varPtr(tt.t_int_1)    '  2

  '
  ' An object seems to be just a pointer, thus t_obj_2 is 4 bytes
  ' apart from t_obj_1:
  '
    debug.print "t_obj_2: " & varPtr(tt.t_obj_2) - varPtr(tt.t_obj_1)    '  4

  '
  ' A TT occupies 24 bytes in memory. This is 7 times 4 bytes:
  '   - t_lng_1
  '   - t_byt_1 and t_byt_2
  '   - t_lng_2
  '   - t_byt_3
  '   - t_int_1
  '   - t_obj_1
  '   - t_obj_2
  '
    debug.print "lenB(T): " & lenB(tt)                                   ' 24

    dim uu as U

  ' debug.print varPtr(uu)

  '
  ' The following statement prints 0.
  ' Again, the address of the type-instance is the
  ' same as the address of its first element, even
  ' if this element happens to be an instance of a type.
  '
    debug.print "u_t___1: " & varPtr(uu.u_t___1 ) - varPtr(uu    )       '  0

  '
  ' The following lines prints 24. This corresponds to the
  ' size that a instance of T occupies in memory:
  '
    debug.print "u_lng_1: " & varPtr(uu.u_lng_1) - varPtr(uu.u_t___1)    ' 24

end sub ' }
