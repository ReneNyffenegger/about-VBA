option explicit


declare sub      GetMem1          lib "msvbvm60.dll" (byVal addr as longPtr,       retVal As byte    )
declare sub      GetMem2          lib "msvbvm60.dll" (byVal addr as longPtr,       retVal As integer )
declare sub      GetMem4          lib "msvbvm60.dll" (byVal addr as longPtr,       retVal As long    )
declare sub      GetMem8          lib "msvbvm60.dll" (byVal addr as longPtr,       retVal As currency)

declare sub      PutMem1          lib "msvbvm60.dll" (byVal addr as longPtr, byVal newVal As byte    )
declare sub      PutMem2          lib "msvbvm60.dll" (byVal addr as longPtr, byVal newVal As integer )
declare sub      PutMem4          lib "msvbvm60.dll" (byVal addr as longPtr, byVal newVal As long    )
declare sub      PutMem8          lib "msvbvm60.dll" (byVal addr as longPtr, byVal newVal As currency)

declare function vbaVarAdd        lib "msvbvm60.dll" alias "__vbaVarAdd" (var1 as variant, var2 as variant) as variant
declare Function vbaVarSub        lib "msvbvm60.dll" alias "__vbaVarSub" (var1 as variant, var2 as variant) as variant
declare function vbaVarMul        lib "msvbvm60.dll" alias "__vbaVarMul" (var1 as variant, var2 as variant) as variant
declare function vbaVarCat        lib "msvbvm60.dll" alias "__vbaVarCat" (var1 as variant, var2 as variant) as variant

declare function vbaObjSetAddref  lib "msvbvm60.dll" alias "__vbaObjSetAddref" (dstObject as any, byVal srcObjPtr as long) as long
declare function vbaObjSet        lib "msvbvm60.dll" alias "__vbaObjSet"       (dstObject as any, byVal srcObjPtr as long) as long

function GetMem1_(byVal addr as longPtr) as byte ' {
    GetMem1 addr, GetMem1_
end function ' }
