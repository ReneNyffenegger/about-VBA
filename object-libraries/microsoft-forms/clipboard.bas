option explicit

'
'  Needs the "Microsoft Forms 2.0 Object Library" reference.
'  If this reference does not show up, it can be
'  added through the file
'    C:\Windows\System32\FM20.dll  or
'    C:\WINDOWS\SysWOW64\FM20.DLL  respectively
'

sub main()

  dim dObj as new dataObject
  dim txt as string

  txt = "This was copied to the clipboard using VBA!"

  dObj.setText "This text comes from a VBA function!"
  dObj.putInClipboard
  
  msgBox "Text was put in clipboard." & vbCrLf & "Check and put something into the clipboard yourself."
  dObj.getFromClipboard
  
  msgBox "Text currently in clipboard is:" & vbCrLf & dObj.getText

end sub
