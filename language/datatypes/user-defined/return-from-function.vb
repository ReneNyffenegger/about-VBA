option explicit


private type XYZ ' {
    num as long
    txt as string
end type ' }


sub main ' {
    dim tp as XYZ

    tp = f

    debug.print "tp.num = " & tp.num
    debug.print "tp.txt = " & tp.txt
end sub ' }


function f as XYZ ' {
    f.num =  42
    f.txt = "Hello World"
end function ' }
