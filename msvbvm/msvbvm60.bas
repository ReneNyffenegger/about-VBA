option explicit

declare sub GetMem1 Lib "msvbvm60.dll" (byVal addr as long,       retVal As byte    )
declare sub GetMem2 Lib "msvbvm60.dll" (byVal addr as long,       retVal As integer )
declare sub GetMem4 Lib "msvbvm60.dll" (byVal addr as long,       retVal As long    )
declare sub GetMem8 Lib "msvbvm60.dll" (byVal addr as long,       retVal As currency)

declare sub PutMem1 Lib "msvbvm60.dll" (byVal addr as long, byVal newVal As byte    )
declare sub PutMem2 Lib "msvbvm60.dll" (byVal addr as long, byVal newVal As integer )
declare sub PutMem4 Lib "msvbvm60.dll" (byVal addr as long, byVal newVal As long    )
declare sub PutMem8 Lib "msvbvm60.dll" (byVal addr as long, byVal newVal As currency)
