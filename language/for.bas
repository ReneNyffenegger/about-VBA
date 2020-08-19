' runVBAFilesInOffice.vbs -word %CD%\for -c main

sub main()

  for i = 5 to 10
    selection.typeText i & " "
  next i

  selection.typeText chr(13)

  for i = 20 to 100 step 15
    selection.typeText i & " "
  next i

  selection.typeText chr(13)

  for i = 1000 to 500 step -125
    selection.typeText i & " "
  next i

  activeDocument.saved = true

end sub
