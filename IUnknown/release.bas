'
'      https://github.com/kumatti1/COM_Test/blob/master/Module1.bas
'

option  explicit
declare ptrSafe sub RtlMoveMemory lib "kernel32" (dest as any, src as any, byVal length as longPtr)

private pVtbl        as longPtr
private func(0 to 2) as longPtr
private iUnk         as IUnknown


private function release(this as longPtr) as long ' {
    msgBox "Exiting Excel"
end function ' }

sub main() ' {


 '  if iUnk is nothing then
 '     debug.print "is nothing"
 '  end if

    debug.print "iUnk:  " & varPtr(iUnk)

    func(2) = vba.int(addressOf release)

    pVtbl  = varPtr(func(0))

    RtlMoveMemory iUnk, varPtr(pVtbl), 4

end sub ' }
