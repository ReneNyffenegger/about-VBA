sub ifTest()

    for i = 1 to 10
    
        if     i > 7 then
               debug.print "i > 7"
        elseif i < 4 then
               debug.print "i < 4"
        else
               debug.print "4 <= i <= 7"
        end if
        
    next i

end sub