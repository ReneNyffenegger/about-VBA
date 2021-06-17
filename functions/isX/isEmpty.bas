sub testIsEmpty()

    dim varVariant as variant
    dim varLong    as long
    dim varObject  as object

    if isEmpty(varVariant) then
       debug.print "varVariant is empty"
    else
       debug.print "varVariant is not empty"
    end if

    if isEmpty(varLong) then
       debug.print "varLong    is empty"
    else
       debug.print "varLong    is not empty"
    end if

    if isEmpty(varObject) then
       debug.print "varObject  is empty"
    else
       debug.print "varObject  is not empty"
    end if

end sub
