option explicit

sub main() ' {

  on error goto failed

    debug.print("42/6 = " & divide(42, 6))
    debug.print("55/0 = " & divide(55, 0))
    debug.print("13/9 = " & divide(13, 9))

  failed:
    debug.print("Failed: " & err.description)

end sub ' }

function divide(a as double, b as double) as double ' {

    if b = 0 then
       call err.raise(1000, "divide", "b must not be zero")
    end if

    divide = a / b

end function ' }
