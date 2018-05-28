' c:\lib\runVBAFilesInOffice\runVBAFilesInOffice.vbs -excel  %CD%\replace-sql-comments -c main
sub main()

  dim sqlText as     string
  dim re      as new regExp

  sqlText =           "select "                      & chr(10) & chr(13)
  sqlText = sqlText & "  *    -- This is a comment " & chr(10) 
  sqlText = sqlText & "from   -- another comment"    & chr(10) & chr(13)
  sqlText = sqlText & "  table-- last comment"

  re.global    = true
  re.multiLine = true
  re.pattern   = "--.*$"

  columns(2).columnWidth = 100
  columns(2).font.name = "Courier New"

  cells(2,2).value = sqlText

  sqlText = re.replace(sqlText, "")

  cells(4,2).value = sqlText


  activeWorkbook.saved = true

end sub
