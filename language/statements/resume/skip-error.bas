option explicit

sub main () ' {
   skipDivisionError
end sub ' }

sub skipDivisionError() ' {
  on error goto err_

    dim res(-3 to 2) as double

    dim i as long
    for i = -3 to 3
       res(i) = 5 / i
       debug.print i
    next i

    exit sub

  err_:

    if err.number = 11 then ' 11: Division by zero
       debug.print("Skipping division by zero error")
       resume next
    end if

    debug.print ("Not skipping " & err.description & " " & err.number)

end sub ' }
