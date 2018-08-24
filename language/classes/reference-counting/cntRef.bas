option explicit

private declare sub RtlMoveMemory lib "kernel32" alias "RtlMoveMemory" ( _
   dest  as     any    , _
   src   as     any    , _
   byVal nbytes as long)


function objRefCnt(obj as IUnknown) as long

    if not obj is nothing then
       RtlMoveMemory objRefCnt, byVal(objPtr(obj)) + 4, 4
       objRefCnt = objRefCnt - 2
    else
       objRefCnt  = 0
    end if

end function
