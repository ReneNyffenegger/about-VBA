sub if_with_or_or_and()

    for i = 1 to 3
    for j = 1 to 3

        if  i=2 or j=2  then
               debug.print "At least on of i or j is 2"
        end if

        if  i=2 and j=2  then
            debug.print "Both, i and j are 2"
        end if

    next j
    next i

end sub
