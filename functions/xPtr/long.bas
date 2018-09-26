option explicit

sub main() ' {

    dim long_1       as long
    dim long_2       as long

    dim addr_1       as longPtr
    dim addr_2       as longPtr

    dim val_addr_1   as longPtr
    dim val_addr_2   as longPtr

    long_1 = 12345678
    long_2 =       42

  '
  ' Determine the addresses of long_1 and long_2:
  '
    addr_1 = varPtr(long_1)
    addr_2 = varPtr(long_2)

  '
  ' The addresses are 4 bytes apart and decreasing. Apparently,
  ' they're allocated on the stack:
  '
    debug.print "addr_1 = " & addr_1 ' 2093616
    debug.print "addr_2 = " & addr_2 ' 2093612

  '
  ' Get the value in memory that the addresses of the
  ' variables point at:
  '
    val_addr_1 = GetMem4_(addr_1)
    val_addr_2 = GetMem4_(addr_2)

  '
  ' Not surprisingly, these values correspond to the
  ' values that were assigned to the variables:
  '
    debug.print "val_addr_1 = " & val_addr_1 ' 12345678
    debug.print "val_addr_2 = " & val_addr_2 '       42

  '
  ' Manipulate the memory that the addresses point at:
  '
    PutMem4 addr_1, 123
    PutMem4 addr_2, 456

  '
  ' The values of the variables were changed:
  '
    debug.print "long_1 = " & long_1 ' 123
    debug.print "long_2 = " & long_2 ' 456

end sub ' }
