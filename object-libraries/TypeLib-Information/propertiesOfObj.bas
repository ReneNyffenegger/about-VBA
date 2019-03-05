'
'   call application.VBE.activeVBProject.references.addFromGuid("{8B217740-717D-11CE-AB5B-D41203C10000}", 1, 0)
'
option explicit

sub propertiesOfObj(byval obj as object) ' {

    dim tlApp  as new tli.tliApplication
    dim tlInfo as     tli.typeInfo

    dim attributes() as string
    dim ix           as long
    dim nofAttrs     as long

    set tlInfo = tlApp.interfaceInfoFromObject(obj)

    nofAttrs = tlInfo.attributeStrings(attributes)

    debug.print "Name             = " & tlInfo.name
    debug.print "GUID             = " & tlInfo.GUID
    debug.print "Kind             = " & tlInfo.typeKindString
    debug.print "AttributeMask    = " & tlInfo.attributeMask

    for ix = lBound(attributes) to uBound(attributes)
        debug.print "                   " & attributes(ix)
    next ix

    debug.print "nof Interfaces   = " & tlInfo.interfaces.count
    debug.print "-----------------------"

    dim mbrInfo as tli.memberInfo

    dim i  as long
    for each mbrInfo In tlInfo.members ' {
        i = i + 1
        debug.print lpad(mbrInfo.memberId, 11) & " " &tlMemberKind(mbrInfo) & rpad(mbrInfo.name, 40)  & ": " & tlTypeName(mbrInfo.returnType)

        dim parInfo as tli.parameterInfo
        for each parInfo in mbrInfo.parameters ' {

            debug.print "   " & tlParamKind(parInfo) & " " & rpad(parInfo.name, 40) & ": " & tlTypeName(parInfo.varTypeinfo)

        next parInfo ' }

'       debug.print "   " & callingConvention(mbrInfo.callConv)
'       if i > 20 then exit sub

        debug.print ""

    next mbrInfo ' }

end sub ' }

private function tlMemberKind(mbr as tli.memberInfo) as string ' {

    select case mbr.descKind ' {
    case   tli.desckind_funcdesc
           select case mbr.invokeKind
           case   tli.invoke_func ' {

                  if mbr.returnType.VarType = VT_VOID then
                     tlMemberKind =  "sub              "
                  else
                     tlMemberKind =  "function         "
                  end if


           case   tli.invoke_propertyGet : tlMemberKind = "property get     "
           case   tli.invoke_propertyPut : tlMemberKind = "property put     "
           case   else                   : tlMemberKind = "?                "
           end    select ' }

    case   tli.desckind_vardesc  : tlMemberKind = "variable     "
    case   tli.desckind_none     : tlMemberKind = "             "
    case   else                  : tlMemberKind = "?            "
    end    select ' }

end function ' }

private function tlParamKind(par as tli.parameterInfo) as string ' {

    if par.flags and paramFlag_fOpt then ' {
       tlParamKind = "optional "
    else
       tlParamKind = ".        "
    end if ' }

    if par.flags and paramFlag_fOut then ' {
       tlParamKind = tlParamKind & "byRef "
    else
       tlParamKind = tlParamKind & "byVal "
    end if ' }

end function ' }

private function tlTypeName(var as tli.varTypeInfo) as string ' {
    dim vType As TliVarType

    vType = var.varType

    if vType and VT_ARRAY Then
       tlTypeName = "()"
       vType = vType and not VT_ARRAY
    end if

    select case vType
    case tli.tliVarType.VT_BOOL      : tlTypeName = tlTypeName & "boolean"
    case tli.tliVarType.VT_BSTR      : tlTypeName = tlTypeName & "string"
    case tli.tliVarType.VT_CY        : tlTypeName = tlTypeName & "currency"
    case tli.tliVarType.VT_DATE      : tlTypeName = tlTypeName & "date"
    case tli.tliVarType.VT_DISPATCH  : tlTypeName = tlTypeName & "object"
    case tli.tliVarType.VT_HRESULT   : tlTypeName = tlTypeName & "HRESULT"
    case tli.tliVarType.VT_I2        : tlTypeName = tlTypeName & "integer"
    case tli.tliVarType.VT_I4        : tlTypeName = tlTypeName & "long"
    case tli.tliVarType.VT_R4        : tlTypeName = tlTypeName & "single"
    case tli.tliVarType.VT_R8        : tlTypeName = tlTypeName & "double"
    case tli.tliVarType.VT_UI1       : tlTypeName = tlTypeName & "byte"
    case tli.tliVarType.VT_UNKNOWN   : tlTypeName = tlTypeName & "IUnknown"
    case tli.tliVarType.VT_VARIANT   : tlTypeName = tlTypeName & "variant"
    case tli.tliVarType.VT_VOID      : tlTypeName = tlTypeName & "any"

    case TliVarType.VT_EMPTY ' {

        if Not var.typeInfo is nothing then
            tlTypeName = tlTypeName & var.typeInfo.name
        end if

    end select ' }

end Function ' }

private function callingConvention(cc as long) as string ' {

    select case cc
    case   tli.cc_cdecl   : callingConvention = "cdecl   "
    case   tli.cc_fastcall: callingConvention = "fastcall"
    case   tli.cc_stdcall : callingConvention = "stdcall "
    case   tli.cc_syscall : callingConvention = "syscall "
    case   else           : callingConvention = "TODO: implement me!"
    end    select

end function ' }

function rpad(text as String, length as integer, optional padChar as string = " ") ' {
 '
 '   https://renenyffenegger.ch/notes/development/languages/VBA/modules/Common/Text
 '
    rpad = text & string(length - len(text), padChar)
end function ' }

function lpad(text as String, length as integer, optional padChar as string = " ") ' {
    lpad = string(length - len(text), padChar) & text
end function ' }
