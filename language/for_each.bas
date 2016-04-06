' runVBAFilesInOffice.vbs -word %CD%\for_each -c main

sub main()

  for each itemain array("foo", "bar", "baz")
    selection.typeText item & chr(13)
  next item

  activeDocument.saved = true

end sub

