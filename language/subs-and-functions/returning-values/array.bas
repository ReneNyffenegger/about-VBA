option explicit

function getArray() as string()

    dim ret(3) as string
	
	ret(1) = "one"
	ret(2) = "two"
	ret(3) = "three"
	
	getArray = ret

end function

sub main()

    dim ary() as string
	
    ary = getArray()
	
	for i = 1 to uBound(ary)
	    msgBox(ary(i))
	next i

end sub