option explicit

sub main ' {
    divideByZero
end sub ' }

sub divideByZero ' {
   on error resume next

   dim num as long
   dim res as long
   for num = -3 to 3
       res = 5 / num
       debug.print("res = " & res & ", err.number = " & err.number & ", err.description = " & err.description)
   next num

end sub ' }
