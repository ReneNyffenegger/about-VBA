option explicit

#if VBA7 then
    public declare ptrSafe sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) ' 64-Bit versions of Excel
#else
    public declare         sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long   ) ' 32-Bit versions of Excel
#end if

sub main()

    const fileName = "c:\temp\sendKeys.txt"

    dim taskID as integer
    taskID = shell("notepad.exe", vbNormalFocus)

  '
  ' Delete final file if it was already written
  '
    if dir(fileName) <> "" then kill fileName

    sendKeys "Hello World!"                , true
    sendKeys "~The tilde means a new line.", true

  '
  ' Bring up the file save as dialog using
  ' the ctrl-s key combination
  '
    sendKeys "^s", true 

  '
  ' wait 0.2 seconds, hopefully, the safe as dialog will be initialized by then.
  '
    sleep 200  ' 

  '
  ' Apparently, this space (or any other character?) is needed to
  ' somehow activate the safe as dialog.
  '
    sendKeys " ", true 
                        
  '
  ' Send the filename
  '
    sendKeys fileName, true

  '
  ' Send enter to close the dialog and
  ' save the file.
  '
    sendKeys "~", true

end sub
