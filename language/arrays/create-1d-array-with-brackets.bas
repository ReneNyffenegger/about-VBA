option explicit

sub main() ' {

   dim ary as variant

   ary = [{ 1, 1, 2, 3, 5, 8 }]

   dim i as integer
   for i = lBound(ary) to uBound(ary)
       debug.print ary(i)
   next i

end sub ' }
