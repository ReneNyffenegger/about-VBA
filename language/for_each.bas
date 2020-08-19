' runVBAFilesInOffice.vbs -word %CD%\for_each -c main

option explicit

sub main()

  dim item as variant
  for each item in array("foo", "bar", "baz")
    selection.typeText item & chr(13)
  next item

  activeDocument.saved = true

end sub
