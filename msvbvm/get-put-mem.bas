option explicit

sub main()
    dim vLong  as long
    dim ptr    as longPtr

    vLong =  3&               +  _
             5& * 256         +  _
             7& * 256*256     +  _
            11& * 256*256*256

    ptr = varPtr(vLong)

    printBytes ptr, 4

    PutMem1 ptr + 0, 0
    PutMem1 ptr + 1, 1
    PutMem1 ptr + 2, 0
    PutMem1 ptr + 3, 0

    debug.print vLong ' 256

end sub



sub printBytes(startAddr as longPtr, cnt as long) ' {

   dim addr as longPtr
   dim b    as byte

   for addr = startAddr to startAddr + cnt - 1
       call  GetMem1(addr, b)
       debug.print b

   next addr

end sub ' }
