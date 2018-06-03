option explicit

sub main()

    dim num as integer

    for num = 0 to 17

        if num and 4 then
           debug.print format$(num, "@@") & " : bit 4 set"
        else
           debug.print format$(num, "@@") & " : bit 4 not set"
        end if

    next num

end sub
