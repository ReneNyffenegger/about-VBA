option explicit

declare sub GetMem1 lib "msvbvm60.dll" (byVal addr as longPtr,       retVal As byte    )
declare sub GetMem2 lib "msvbvm60.dll" (byVal addr as longPtr,       retVal As integer )
declare sub GetMem4 lib "msvbvm60.dll" (byVal addr as longPtr,       retVal As long    )
declare sub GetMem8 lib "msvbvm60.dll" (byVal addr as longPtr,       retVal As currency)

declare sub PutMem1 lib "msvbvm60.dll" (byVal addr as longPtr, byVal newVal As byte    )
declare sub PutMem2 lib "msvbvm60.dll" (byVal addr as longPtr, byVal newVal As integer )
declare sub PutMem4 lib "msvbvm60.dll" (byVal addr as longPtr, byVal newVal As long    )
declare sub PutMem8 lib "msvbvm60.dll" (byVal addr as longPtr, byVal newVal As currency)


function GetMem1_(byVal addr as longPtr) as byte ' {
    GetMem1 addr, GetMem1_
end function ' }
