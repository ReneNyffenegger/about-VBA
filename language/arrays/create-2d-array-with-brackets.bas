option explicit

sub main()

   dim ary_2d as variant

   ary_2d = [ { 1,2,3 ; "foo","bar","baz" }]

   dim i as integer
   for i = 1 to 3
       debug.print ary_2d(1, i) & ": " & ary_2d(2, i)
   next i

end sub
