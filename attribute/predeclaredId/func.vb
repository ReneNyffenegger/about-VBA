option explicit

sub main() ' {

    predeclared.setMemberValue 42

    dim abc as new predeclared
    abc.setMemberValue 99

    debug.print "member value of predeclared is: " & predeclared.getMemberValue
    debug.print "member value of abc is:         " &         abc.getMemberValue

end sub ' }
