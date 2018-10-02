option explicit


declare sub      GetMem1          lib "msvbvm60.dll" (byVal addr as longPtr,       retVal As byte    )
declare sub      GetMem2          lib "msvbvm60.dll" (byVal addr as longPtr,       retVal As integer )
declare sub      GetMem4          lib "msvbvm60.dll" (byVal addr as longPtr,       retVal As long    )
declare sub      GetMem8          lib "msvbvm60.dll" (byVal addr as longPtr,       retVal As currency)

declare sub      PutMem1          lib "msvbvm60.dll" (byVal addr as longPtr, byVal newVal As byte    )
declare sub      PutMem2          lib "msvbvm60.dll" (byVal addr as longPtr, byVal newVal As integer )
declare sub      PutMem4          lib "msvbvm60.dll" (byVal addr as longPtr, byVal newVal As long    )
declare sub      PutMem8          lib "msvbvm60.dll" (byVal addr as longPtr, byVal newVal As currency)

' declare function VarPtr           lib "msvbvm60.dll" (byVal ptr  as long) as long

declare function vbaVarAdd        lib "msvbvm60.dll" alias "__vbaVarAdd" (var1 as variant, var2 as variant) as variant
declare Function vbaVarSub        lib "msvbvm60.dll" alias "__vbaVarSub" (var1 as variant, var2 as variant) as variant
declare function vbaVarMul        lib "msvbvm60.dll" alias "__vbaVarMul" (var1 as variant, var2 as variant) as variant
declare function vbaVarCat        lib "msvbvm60.dll" alias "__vbaVarCat" (var1 as variant, var2 as variant) as variant

' vbaCopyBytes* {
'
'       vbaCopyBytes: similar to RtlMoveMemory although for non-overlapping memory only.
declare function vbaCopyBytes     lib "msvbvm60.dll" alias "__vbaCopyBytes"     (byVal length as long, dst as any, src as any) as long
'
'       vbaCopyBytesZero: Might be __vbaCopyBytes, followed by zeroing out the source memory block with zeroes.
'                         But what do I know?
declare function vbaCopyBytesZero lib "msvbvm60.dll" alias "__vbaCopyBytesZero" (byVal length as long, dst as any, src as any) as long
' }

declare function vbaObjSetAddref  lib "msvbvm60.dll" alias "__vbaObjSetAddref"  (dstObject as any, byVal srcObjPtr as long) as long
declare function vbaObjSet        lib "msvbvm60.dll" alias "__vbaObjSet"        (dstObject as any, byVal srcObjPtr as long) as long

function GetMem1_(byVal addr as longPtr) as byte ' {
    GetMem1 addr, GetMem1_
end function ' }

function GetMem2_(byVal addr as longPtr) as integer ' {
    GetMem2 addr, GetMem2_
end function ' }

function GetMem4_(byVal addr as longPtr) as long ' {
    GetMem4 addr, GetMem4_
end function ' }

function GetMem8_(byVal addr as longPtr) as currency ' {
    GetMem8 addr, GetMem8_
end function ' }
