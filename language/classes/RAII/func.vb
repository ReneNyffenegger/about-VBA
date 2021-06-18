option explicit

sub main()
    useResource
end sub


sub useResource()

    dim r as resource : set r = new resource

    debug.print "try A" : if A then exit sub
    debug.print "try B" : if B then exit sub
    debug.print "try C" : if C then exit sub
    debug.print "try D" : if D then exit sub

end sub

function A as boolean
    A = false
end function

function B as boolean
    B = false
end function

function C as boolean
    C = true
end function

function D as boolean
    D = false
end function

