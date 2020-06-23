option explicit

' global dirName as string

sub main() ' {

    dim dirName as string
    dirName = environ$("TEMP") & "\VBA-Test"

    createFiles dirName

    showFilesInDirectory dirName

    kill dirName & "\four.txt"
    showFilesInDirectory dirName

  '
  ' Delete files with wildcard:
  '
    kill dirName & "\t*.txt"
    showFilesInDirectory dirName

    deleteFileIfExists dirName & "\one.txt"
    deleteFileIfExists dirName & "\one.txt"

  '
  ' Remove last remaining file so as to
  ' be able to delete directory:
  '

    kill dirName & "\five.txt"

    rmDir dirName

end sub ' }

sub createFiles(dirName as string) ' {

    mkDir dirName

    createFile dirName & "\one.txt"
    createFile dirName & "\two.txt"
    createFile dirName & "\three.txt"
    createFile dirName & "\four.txt"
    createFile dirName & "\five.txt"

end sub ' }

sub createFile(fileName as string) ' {

    dim fn as integer

    fn = freeFile()
    open fileName for output as #fn

    close# fn

end sub ' }

sub showFilesInDirectory(dirName as string) ' {

    dim fileName as string

'   If Right(sPath, 1) <> "\" Then
'       sPath = sPath & "\"
'   End If

'   If sFilter = "" Then
'       sFilter = "*.*"
'   End If


  '
  ' Apparently, dir(…) requires trailing back slash for
  ' getting file names in direcotry:
  '
    fileName = dir(dirName & "\") ' "

    debug.print "Files in directory now:"
    do until fileName = "" ' {

       debug.print "  " & fileName

     '
     ' Use dir without parameter to obtain next file name:
     '
       fileName = dir
    loop ' }

end sub ' }

sub deleteFileIfExists(fileName as string) ' {

    if dir(fileName) <> "" then
       kill fileName
       debug.print "file " & fileName & " was deleted"
    else
       debug.print "file " & fileName & " does not exist"
    end if

end sub ' }
