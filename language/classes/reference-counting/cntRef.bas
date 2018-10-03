option explicit

private declare sub RtlMoveMemory lib "kernel32" alias "RtlMoveMemory" ( _
   dest  as     any    , _
   src   as     any    , _
   byVal nbytes as long)


function objRefCnt(obj as IUnknown) as long ' {

  ' debug.print "objPtr(obj) = " & objPtr(obj)

    if not obj is nothing then
    '
    '  Apparently, an object's ref Count is pointed at by the
    '  second member in a VB COM object:
    '
       RtlMoveMemory objRefCnt, byVal objPtr(obj) + 4, 4
       objRefCnt = objRefCnt - 2
    else
       objRefCnt  = 0
    end if

end function ' }
