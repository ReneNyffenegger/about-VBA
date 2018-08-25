option explicit

sub main()

    dim ary() as integer
    dim byt   as byte
    dim boo   as boolean
    dim cur   as currency
    dim dat   as date
    dim dbl   as double
    dim igr   as integer
    dim lng   as long
  ' dim lll   as longLong
    dim lPt   as longPtr
    dim obj   as object
    dim sgl   as single
    dim var   as variant
    dim xxx

    debug.print "array:    " & varType(ary)    ' 8194
    debug.print "byte:     " & varType(byt)    ' 17
    debug.print "boolean:  " & varType(boo)    ' 11
    debug.print "currency: " & varType(cur)    ' 6
    debug.print "date:     " & varType(dat)    ' 7
    debug.print "double:   " & varType(dbl)    ' 5
    debug.print "integer:  " & varType(igr)    ' 2
    debug.print "long:     " & varType(lng)    ' 3
  ' debug.print "longLong: " & varType(lll)
    debug.print "longPtr:  " & varType(lPt)    ' 3
    debug.print "object:   " & varType(obj)    ' 9
    debug.print "single:   " & varType(sgl)    ' 4
    debug.print "variant:  " & varType(var)    ' 0
    debug.print "xxx:      " & varType(xxx)    ' 0

    var = 12
    xxx = 49.31
    debug.print "variant:  " & varType(var)    ' 2
    debug.print "xxx:      " & varType(xxx)    ' 5


end sub
