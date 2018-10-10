version 1.0 class
begin
  multiUse                   = -1  ' true
  persistable                =  0  ' false: not persistable
  dataBindingBehavior        =  0  ' vbNone
  dataSourceBehavior         =  0  ' vbNone
  MTSTransactionMode         =  0  ' NotAnMTSObject
end
attribute vb_name            = "EnumVARIANT"
attribute vb_globalNameSpace =  false
attribute vb_creatable       =  true
attribute vb_predeclaredId   =  false
attribute vb_exposed         =  false

option explicit

public interface_IEnumVARIANT as IEnumVARIANT

public function get_IEnumVARIANT() as IEnumVARIANT ' {
       attribute       get_IEnumVARIANT.vb_userMemId   = -4

     ' The attribute VB_UserMemId determines the method's DispatchID.
     ' A DispatchID with a negative number has a special meaning.
     ' In this case, -4 corresponds to DISPID_NEWENUM
     ' See also http://www.papwalker.com/ref101/ccol.html
     '
     ' The returned datatype (or interface) IEnumVARIANT is
     ' predeclared by Visual Basic for Applications.
     '
       set get_IEnumVARIANT = interface_IEnumVARIANT
end function ' }

' vim: ft=vb
