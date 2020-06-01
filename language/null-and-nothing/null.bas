option explicit

sub testEmptyAndNull(var as variant)

    if isEmpty(var) then
       debug.print "| var is empty"
    end if

    if isNull(var) then
       debug.print "| var is null"
    end if

    debug.print "| var = " & var
    debug.print ""

end sub

sub main()

    dim dbl as double
    dim var as variant


  ' dbl = null ' null can only be assigned to variants
  '            ' Otherwise, a Run-time error 94 is thrown: Invalid use of Null

  ' If a variant is null can be tested with isNull:

    testEmptyAndNull var
'
'   | var is empty
'   | var =

    var = null
    testEmptyAndNull var
'
'   | var is null
'   | var =

    var = 42
    testEmptyAndNull var
'
'   | var = 42

end sub
