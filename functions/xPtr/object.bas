option explicit

sub main() ' {

    dim o       as new obj

    dim obj_ptr as longPtr
    dim var_ptr as longPtr

    obj_ptr = objPtr(o)
    var_ptr = varPtr(o)

    debug.print "var_ptr          = " & var_ptr
    debug.print "obj_ptr          = " & obj_ptr
  '
  ' var_ptr points at obj_ptr
  '
    debug.print "GetMem4(var_ptr) = " & GetMem4_(var_ptr)

end sub ' }

' var_ptr          = 2093616
' obj_ptr          = 93661960
' GetMem4(var_ptr) = 93661960
