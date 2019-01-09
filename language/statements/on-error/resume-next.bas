option explicit

sub main ' {
    divideByZero
end sub ' }

sub divideByZero ' {
   on error resume next

   dim i as long
   for i = -3 to 3
       debug.print("5/" & i & " = " & 5/i)
   next i

end sub ' }
