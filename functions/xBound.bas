option explicit

sub main() ' {

  ' declare an array of 6 longs, indexed from 0 to 5
    dim a_0_5(0 to 5) as long

  ' declare another array of 6 longs:
    dim a_3_8(3 to 8) as long

  ' declare another array. Since option base is not
  ' explicitely set, it defaults to 0, so the lower
  ' bound is 0.
     dim a_x_5(5)

    showBounds a_0_5 ' lBound: 0, uBound: 5
  showBounds a_3_8 ' lBound: 3, uBound: 8
  showBounds a_x_5 ' lBound: 0, uBound: 5

end sub ' }

sub showBounds(a as variant) ' {
    debug.print "lBound: " & lBound(a) ", uBound: " & uBound(a)
end sub ' }
