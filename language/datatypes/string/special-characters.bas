option explicit

sub special_characters() ' {

    dim kappa  as string
    dim smiley as string

    kappa  = "κ"
    smiley = "☺"
    debug.print("Smiley: " & smiley & " (" & hex(ascW(smiley)) & "), kappa = " & kappa & " (" & hex(ascW(kappa)) & ")")
    msgBox     ("Smiley: " & smiley & ", kappa = " & kappa)

    kappa  = chrW(&h03BA) ' Unicode character GREEK SMALL LETTER KAPPA
    smiley = chrW(&h263A) ' Unicode character WHITE SMILING FACE
    debug.print("Smiley: " & smiley & " (" & hex(ascW(smiley)) & "), kappa = " & kappa & " (" & hex(ascW(kappa)) & ")")
    msgBox     ("Smiley: " & smiley & ", kappa = " & kappa)

  '
  ' Trying to write a file with the special characters
  '
    open environ("TEMP") & "\special-characters.txt" for output as #1
    print #1, "Kappa:  " & kappa
    print #1, "Smiley: " & smiley
    close #1

  '
  ' Using ADODB to write a file:
  '
    dim adoStream as new adodb.stream
    adoStream.charset = "utf-8"
    adoStream.open
    adoStream.writeText "Kappa:  " & kappa , streamWriteEnum.stWriteLine
    adoStream.writeText "Smiley: " & smiley, streamWriteEnum.stWriteLine
    adoStream.saveToFile environ("TEMP") & "\special_characters.utf8"
    adoStream.close

end sub ' }
