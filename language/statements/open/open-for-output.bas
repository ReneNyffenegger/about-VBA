sub writeToFile()

  open environ("TEMP") & "\vba-file-test.txt" for output as #1

  print #1, "Hello world"
  print #1, "Second line"

  close #1

end sub
