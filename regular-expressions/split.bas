' runVBAFilesInOffice.vbs -word %CD%\split -c main

sub main() ' {

' http://stackoverflow.com/a/35254567/180275

  dim text   as string
  dim nums() as string

  text = "twenty-two22fourty-nine49six6eleven11five hundred500"

  nums = regexpSplit(text, "\d+")

  selection.font.name = "Courier New"

  for each num in nums
      selection.typeText ">" & num & "<" & vbCrLf
  next num

  activeDocument.saved = true
 
end sub ' }

private function regexpSplit(text as string, pattern as string) as string() ' {

  dim text_0 as string
  dim re     as new regExp

  re.pattern   = pattern
  re.global    = true
  re.multiLine = true

  text_0 = re.replace(text, vbNullChar)

  regexpSplit = strings.split(text_0, vbNullChar)

end function ' }
