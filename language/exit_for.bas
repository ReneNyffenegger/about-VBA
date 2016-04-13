' runVBAFilesInOffice.vbs -word %CD%\exit_for -c main

sub main()

  dim i, r as long
  for i = 1 to 10

    r = i

    if i = 7 then
       exit for
    end if

  next i

  selection.typeText("r = " & r) 

  activeDocument.saved = true

end sub

