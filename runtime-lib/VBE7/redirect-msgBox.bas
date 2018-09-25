'
' Idea found at https://gist.github.com/kumatti1/f9ec1a011987dd3f2753
'

option explicit


dim newBytes     (0 to 5) as byte
dim originalBytes(0 to 5) as byte

dim funcPtr   as long

public sub undoReplacement()
    RtlMoveMemory byVal funcPtr, byVal VarPtr(originalBytes(0)), 6
end sub

public sub redirect_msgBox(p as long) ' {

    dim flOldProtect     as long
    dim vbe7_dll         as long

    vbe7_dll  = GetModuleHandle("vbe7.dll")
    funcPtr   = GetProcAddress(vbe7_dll, "rtcMsgBox")

    if VirtualProtect(byVal funcPtr, 6, PAGE_EXECUTE_RW, flOldProtect) <> 0 then ' {

       RtlMoveMemory byVal VarPtr(originalBytes(0)), byVal funcPtr, 6


       newBytes(0) = &h68 ' Opcode for PUSH
    '
    '  4 byte value (address of new function) to be PUSHed
    '
       RtlMoveMemory byVal VarPtr(newBytes(1)), byVal VarPtr(p), 4

       newBytes(5) = &hC3 ' Opcode FOR RET

       RtlMoveMemory byVal funcPtr, byVal VarPtr(newBytes(0)), 6

    end if ' }

end sub ' }

private function replacementFunc(prompt, byVal buttons&, title, helpFile, context) as long ' {
  '
  ' This function will be called instead of the »real« msgBox after the
  ' redirection is put in operation.
  '

    debug.print prompt

  '
  ' Instead of calling debug.print, the replacement function might
  ' also call MessageBox with a different title...:
  '
  ' replacementFunc = MessageBox(Application.hwnd, strPtr(prompt), strPtr("NOTE THE CHANGED TITLE!!!"), 0)
  '
end function ' }

sub main() ' {

  '
  ' The first call of msgBox uses VBA's default msgBox:
  '
    msgBox "Before redirect"

    call redirect_msgBox(vba.cLng(addressOf replacementFunc))

  '
  ' After setting up the redirection, calling msgBox is routed
  ' to debug.print. Thus, no message box is displayed here, instead
  ' a message is written to the debug.out destination:
  '
    msgBox "After redirect"

    undoReplacement

  '
  ' After undoing the replacement, everything works as usual again:
  '
    msgBox "finishing"

end sub ' }
