' runVBAFilesInOffice.vbs -word %CD%\on_error_goto -c main
'
' based upon http://stackoverflow.com/a/6050733/180275

sub main()

  selection.font.name = "Courier New"

  selection.typeText "main" & chr(13)

  selection.typeText "-> A()" & chr(13)

  call A()

  activeDocument.saved = true

end sub

sub A()

  selection.typeText "  A" & chr(13)

  selection.typeText "  -> B()" & chr(13)

  call B()

end sub

sub B()

  selection.typeText "    B" & chr(13)

  selection.typeText "    -> C()" & chr(13)

  call C()

end sub

sub C()

  on error goto err_C

  selection.typeText "      C"  & chr(13)

  selection.typeText "      -> D()"  & chr(13)

  call D()


 done_C:

  exit sub

 err_C:
  
  msgBox "something's wrong" & vbCrLf & err.description & " [" & err.number & "]"

  resume done_C

' Note this second resume.
' It will never execute in normal processing, since the Resume <label>
' statement will send the execution elsewhere. It can be a godsend for debugging,
' though. When you get an error notification, choose Debug (or press Ctl-Break,
' then choose Debug when you get the "Execution was interrupted" message). The
' next (highlighted) statement will be either the MsgBox or the following
' statement. Use "Set Next Statement" (Ctl-F9) to highlight the bare Resume, then
' press F8. This will show you exactly where the error was thrown.
  resume

end sub

sub D()

  dim a, b, c as integer

  selection.typeText "        D"  & chr(13)
  selection.typeText "        ."  & chr(13)

  a = 5 - 3

  selection.typeText "        ."  & chr(13)

  b = a - 2

  selection.typeText "        ."  & chr(13)

  c = a / b ' Boom!

end sub
