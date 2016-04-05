' runVBAFilesInOffice.vbs -word %CD%\select -c main

sub main()

  selection.font.name = "Courier New"

  selection.typeText "1: " & num_to_text(1) & chr(13)
  selection.typeText "2: " & num_to_text(2) & chr(13)
  selection.typeText "3: " & num_to_text(3) & chr(13)

  activeDocument.saved = true

end sub

private function num_to_text(i as long) ' {

  select case i
  case 1 : num_to_text = "one"
  case 2 : num_to_text = "two"
  case 3 : num_to_text = "three"
  case 4 : num_to_text = "four"
  case 5 : num_to_text = "five"
  case 6 : num_to_text = "six"
  case 7 : num_to_text = "seven"
  case 8 : num_to_text = "eight"
  case 9 : num_to_text = "nine"
  end select

end function ' }
