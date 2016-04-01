' runVBAFilesInOffice.vbs -word %CD%\literals -c main

sub main()


  dim dt_1, dt_2, dt_3 as date

  dt_1  = #08/28/2010#
  dt_2  = "August 28, 2011"
  dt_3  = #08/28/2011 11:22:33 AM#

  
  selection.font.name = "Courier New"

  selection.typeText dt_1 & chr(13)
  selection.typeText dt_2 & chr(13)
  selection.typeText dt_3 & chr(13)

  activeDocument.saved = true

end sub
