attribute vb_predeclaredId   = true
' attribute vb_globalNameSpace = true

option explicit

private member as long

sub class_initialize()
    debug.print "An instance of predeclared is being initialized"
end sub

sub class_terminate()
    debug.print "terminate predeclared instance, member = " & member
end sub

sub setMemberValue(m as long) ' {
    debug.print "setting member value to " & m
    member = m
end sub ' }

function getMemberValue() as long' {
    getMemberValue = member
end function ' }
