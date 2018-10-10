attribute vb_name = "EnumVARIANT_module"
'
' Using some ideas of Dexter Freivald which were really helpful.
'
option explicit


type GUID_as_16_bytes ' {
     b_00 as byte: b_01 as byte: b_02 as byte: b_03 as byte
     b_04 as byte: b_05 as byte
     b_06 as byte: b_07 as byte
     b_08 as byte: b_09 as byte
     b_0a as byte: b_0b as byte: b_0c as byte: b_0d as byte
     b_0e as byte
     b_0f as byte
end type ' }

type IUnknown_vtbl_T ' {
     QueryInterface  as longPtr
     AddRef          as longPtr
     Release         as longPtr
end  type ' }

type g_IEnumVariant_vtbl_T ' {
   '
   '  IEnumVARIANT inherits from IUnknown, thus the first element
   '  of g_IEnumVariant_vtbl_T needs to be an IUnknown_vtbl_T:
   '
      IUnknown_vtbl  as IUnknown_vtbl_T
   '
   '  The four function pointers to the specific methods of IEnumVARIANT:
   '
      Next           as longPtr
      Skip           as longPtr
      Reset          as longPtr
      Clone          as longPtr

end  type ' }

'
' The function pointer table for the IEnumVARIANT interface.
' Only one is needed, therefor, it is declared globally.
'
private g_IEnumVariant_vtbl as g_IEnumVariant_vtbl_T

private type TENUMERATOR ' {
    vTablePtr          as longPtr
    refCount           as long
    curItem            as long       ' The index of the next item to be returned.
    items              as variant    ' The array that stores the elements over which it will be iterated.
end type ' }

const NULL_ = 0

'
' Apparently, the declaration of GetMem4 in msvbvm60.bas is not compatible with the following declaration:
' It is needed for the function DLng (which seems not to be called anyway, but what do I know?)
'
private declare sub GetMem4_local lib "msvbvm60" alias "GetMem4" (byRef src  as any, dest as any)

public function items(paramArray v()) as EnumVARIANT ' {

     dim IEnumVARIANT_implementation as IEnumVARIANT

   '
   ' For some reason, a parramArray values cannot be passed directly to sub that takes
   ' a variant. It can, however, be assigned to a variant which then can be passed
   ' to the sub.
   '
     dim v_ as variant
     v_ = v
     set IEnumVARIANT_implementation =  get_IEnumVariant_vtbl_etc(v_)

     set items = new EnumVARIANT
'    items.setInterface IEnumVARIANT_implementation
     set items.interface_IEnumVARIANT = IEnumVARIANT_implementation

end function ' }

public function get_IEnumVariant_vtbl_etc (items_ as variant) as IEnumVARIANT ' { get_IEnumVariant_vtbl_etc
'
'   The returneed datatype (or interface) IEnumVARIANT is already
'   defined by Visual Basic for Applications.
'

    dim this As TENUMERATOR

    if  g_IEnumVariant_vtbl.IUnknown_vtbl.QueryInterface = NULL_ then ' {
     '
     '  The IEnumVARIANT virtual function table has not yet been initialized.
     '  Do it now:
     '
        g_IEnumVariant_vtbl.IUnknown_vtbl.QueryInterface = vba.cLng( addressOf IUnknown_QueryInterface )
        g_IEnumVariant_vtbl.IUnknown_vtbl.AddRef         = vba.cLng( addressOf IUnknown_AddRef         )
        g_IEnumVariant_vtbl.IUnknown_vtbl.Release        = vba.cLng( addressOf IUnknown_Release        )
      ' ----------------------------------------------------------------------------------------------
        g_IEnumVariant_vtbl.Next                         = vba.cLng( addressOf IEnumVARIANT_Next       )
        g_IEnumVariant_vtbl.Skip                         = vba.cLng( addressOf IEnumVARIANT_Skip       )
        g_IEnumVariant_vtbl.Reset                        = vba.cLng( addressOf IEnumVARIANT_Reset      )
        g_IEnumVariant_vtbl.Clone                        = vba.cLng( addressOf IEnumVARIANT_Clone      )

    end if ' }



    this.vTablePtr    = vba.varPtr(g_IEnumVariant_vtbl)
    this.curItem      = lBound(items_)
    this.refCount     = 1
    this.items        = items_

    dim pThis as longPtr

  '
  ' Allocate "COM" memory
  '
    pThis = CoTaskMemAlloc(lenB(This))
    RtlMoveMemory byVal pThis, this, lenB(this)

  '
  ' This is sort of unbelievable, but "this" must be zeroed out.
  '
  ' Don Box states the reason for this (Advanced Visual Basic 6, p. 149):
  '    VB thinks the data in Struct needs to be  freed when the function goes out of scope VB has no
  '    way of knowing that ownership of the structure has moved elsewhere. If the
  '    structure contains object or variable-size String or array types, VB will
  '    kindly free them for you when the object goes out of scope. But you are still
  '    using the structure, so this is an unwanted favor. To prevent VB from freeing
  '    referenced memory in the stack object, simply ZeroMemory the structure. When
  '    you apply the CopyMemory call's ANSI/UNICODE precautions to ZeroMemory, you
  '    get the transfer code seen in the listing.
  '
  ' Apparently, the combination of RtlMoveMemory and RtlZeroMemory could be
  ' achieved in one go with vbaCopyBytesZero (declared in msvbvm60.dll).
  '
    RtlZeroMemory                             this, lenB(this)

    dim addr_IEnumVariant as longPtr
    dim addr_pThis        as longptr

    addr_IEnumVariant = varPtr(get_IEnumVariant_vtbl_etc)
    addr_pThis        = varPtr(pThis)

    RtlMoveMemory byVal addr_IEnumVariant, byVal addr_pThis, lenB(pThis)

end function ' }

private function is_IID_IUnknown(byRef g as GUID_as_16_bytes) as boolean ' { 00000000-0000-0000-C000-000000000046

    is_IID_IUnknown =  _
       g.b_00 = &h00 and g.b_01 = &h00 and g.b_02 = &h00 and g.b_03 = &h00 and _
       g.b_04 = &h00 and g.b_05 = &h00                                     and _
       g.b_06 = &h00 and g.b_07 = &h00                                     and _
       g.b_08 = &hc0 and g.b_09 = &h00                                     and _
       g.b_0a = &h00 and g.b_0b = &h00 and g.b_0c = &h00 and g.b_0d = &h00 and _
       g.b_0e = &h00 and g.b_0f = &h46

end function ' }

private function is_IID_IEnumVARIANT(g as GUID_as_16_bytes) as boolean ' { 00020404-0000-0000-C000-000000000046

    is_IID_IEnumVARIANT = _
       g.b_00 = &h04 and g.b_01 = &h04 and g.b_02 = &h02 and g.b_03 = &h00 and _
       g.b_04 = &h00 and g.b_05 = &h00                                     and _
       g.b_06 = &h00 and g.b_07 = &h00                                     and _
       g.b_08 = &hc0 and g.b_09 = &h00                                     and _
       g.b_0a = &h00 and g.b_0b = &h00 and g.b_0c = &h00 and g.b_0d = &h00 and _
       g.b_0e = &h00 and g.b_0f = &h46

end function ' }

' IUnknown_QueryInterface {
private function IUnknown_QueryInterface(        _
            byRef this      as TENUMERATOR     , _
                  g         as GUID_as_16_bytes, _
            byVal ppvObject as longPtr           _
                                         ) as long


    if  ppvObject = NULL_ then
        IUnknown_QueryInterface = E_POINTER
        exit function
    end If


    if is_IID_IUnknown(g) or is_IID_IEnumVARIANT(g) then

        dim addr_this      as longPtr
        dim addr_addr_this as longPtr
        dim addr_ppvObject as longPtr

        addr_this      = vba.varPtr(this)
        addr_addr_this = vba.varPtr(addr_this)
        addr_ppvObject = vba.varPtr(ppvObject)

        RtlMoveMemory byVal ppvObject, addr_this, lenB(ppvObject)

        IUnknown_AddRef           this
        IUnknown_QueryInterface = S_OK

    else
     '
     '  The requested interface is not supported, report it
     '  by returning E_NOINTERFACE
     '
        IUnknown_QueryInterface = E_NOINTERFACE
    end If

end function ' }

private function IUnknown_AddRef(byRef this as TENUMERATOR) as long ' {

    this.refCount   = this.refCount + 1
    IUnknown_AddRef = this.refCount

end function ' }

private function IUnknown_Release(byRef this As TENUMERATOR) as long ' {

   this.refCount = this.refCount - 1

   if this.refCount = 0 then ' {
      CoTaskMemFree vba.varPtr(this)
   end if ' }

end function ' }

' { IEnumVARIANT_Next
private function IEnumVARIANT_Next(           _
           byRef this         as TENUMERATOR, _
           byVal celt         as long       , _
           byVal rgVar        as long       , _
           byVal pCeltFetched as long         _
  ) as long

  ' Parameters
  '   celt        : the number of elements to be retrieved
  '   rgVar       : An array of at least size celt in which the elements are to be returned.
  '   pCeltFetched: The number of elements returned in rgVar, or NULL.
  '
  ' Return Value
  '   S_OK        : The number of elements returned is celt.
  '   F_FALSE     : The number of elements returned is less than celt.


    if  rgVar = NULL_ then
        IEnumVARIANT_Next = E_POINTER
        exit function
    end if

    dim fetched as long
    do until this.curItem > uBound(this.items) ' {

          '
          ' VariantCopy is defined in oleaut32.dll
          '
            VariantCopy    rgVar, this.items(this.curItem)

            this.curItem = this.curItem + 1
            fetched      = fetched      + 1

            if fetched  = celt then
               exit do
            end if

            rgVar       = pointerAddition(rgVar, 16&)
    loop ' }


    if pceltFetched then
        debug.print "pceltFetched is unexepctedly true"
        DLng(pceltFetched) = fetched
    end if

    if fetched < celt then
       debug.print "fechted < celt"
       IEnumVARIANT_Next = S_FALSE
    end if

end function ' }

private function IEnumVARIANT_Skip(byRef this as TENUMERATOR, byVal celt as long) as long ' {
    IEnumVARIANT_Skip = E_NOTIMPL
end function ' }

private function IEnumVARIANT_Reset(ByRef This As TENUMERATOR) as long ' {
    IEnumVARIANT_Reset = E_NOTIMPL
end function ' }

private function IEnumVARIANT_Clone(ByRef This As TENUMERATOR, ByVal ppEnum as long) as long ' {
    IEnumVARIANT_Clone = E_NOTIMPL
end function ' }

private function pointerAddition(ByVal pointer as longPtr, byVal offset as longPtr) as longPtr ' {
    pointerAddition = (pointer xor &H80000000) + (offset xor &H80000000)
end function ' }

private property let DLng(byVal Address as long, byVal value As Long) ' {
    GetMem4_local Value, byVal Address
end property ' }
