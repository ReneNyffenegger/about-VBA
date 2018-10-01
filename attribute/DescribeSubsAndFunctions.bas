attribute vb_name = "DescribeSubsAndFunctions"
option explicit

sub thisIsASub()
    attribute thisIsASub.vb_description = "This is the sub's description"
    debug.print "thisIsASub"
end sub

function thisIsAFunction as string
    attribute thisIsAFunction.vb_description = "This is the function's description"
    thisIsAFunction = "foo bar baz"
end function
