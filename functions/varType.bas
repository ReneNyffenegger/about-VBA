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
    dim txt   as string
    dim var   as variant
    dim xxx

    debug.print "array:    " & varType(ary)    ' vbArray + vbInteger = 8192 + 2 = 8194
    debug.print "byte:     " & varType(byt)    ' vbByte              = 17
    debug.print "boolean:  " & varType(boo)    ' vbBoolean           = 11
    debug.print "currency: " & varType(cur)    ' vbCurrency          =  6
    debug.print "date:     " & varType(dat)    ' vbDAte              =  7
    debug.print "double:   " & varType(dbl)    ' vbDouble            =  5
    debug.print "integer:  " & varType(igr)    ' vbInteger           =  2
    debug.print "long:     " & varType(lng)    ' vbLong              =  3
  ' debug.print "longLong: " & varType(lll)
    debug.print "longPtr:  " & varType(lPt)    '                     =  3
    debug.print "object:   " & varType(obj)    ' vbObject            =  9
    debug.print "single:   " & varType(sgl)    ' vbSingle            =  4
    debug.print "string:   " & varType(txt)    ' vbString            =  8
    debug.print "variant:  " & varType(var)    ' vbEmtpy             =  0  -- vbEmpty because no value is assigned
    debug.print "xxx:      " & varType(xxx)    ' vbEmpty             =  0

    var = 12
    xxx = 49.31
    debug.print "variant:  " & varType(var)    ' vbInteger          =   2  -- Note that varType does not return 12 (= vbVariant, but the value of data that the variant stores
    debug.print "xxx:      " & varType(xxx)    ' vbDouble           =   5

end sub
