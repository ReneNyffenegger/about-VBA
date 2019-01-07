option explicit

sub main() ' {

     debug.print ("25/5 = " & divide(25, 5))
     debug.print ("42/0 = " & divide(42, 0))

end sub ' }

function divide(a as double, b as double) as double ' {

  on error goto failed

    divide = a / b

  exit function

  failed:
     debug.print("division failed")
     debug.print("  err.description = " & err.description)
     debug.print("  err.number      = " & err.number     )
     debug.print("  err.source      = " & err.source     )

end function ' }
