sub main()

    test "foo"
    test "foo bar"
    test "foo 999 bar baz"
    test "F"
    test "I"
    test "FOO BAR BAZ"
    test "42"

end sub

sub test(text)

    debug.print "text = " & text
    if text like "*bar*" then
       debug.print "  text contains bar"
    else
       debug.print "  text does not contain bar"
    end if

    if text like "*bar" then
       debug.print "  text ends in bar"
    else
       debug.print "  text does not end in bar"
    end if

    if text like "???" then
       debug.print "  text is three characters long"
    else
       debug.print "  text is not three characters long"
    end if

    if text like "*#*" then
       debug.print "  text contains at least a numeric character (digit)"
    else
       debug.print "  text does not contain a numeric character"
    end if

    if text like "[A-F]" then
       debug.print "  text is an uppercase letter between A and F"
    else
       debug.print "  text is not an uppercase letter between A and F"
    end if

end sub
