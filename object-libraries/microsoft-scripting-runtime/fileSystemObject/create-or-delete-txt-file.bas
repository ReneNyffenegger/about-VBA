option explicit

sub main() ' {

    dim pathToFile as string
    pathToFile = environ$("TEMP") & "\someFile.txt"

    dim fs as new scripting.fileSystemObject

    if fs.fileExists(pathToFile) then ' {

       msgBox(pathToFile & " exists." & vbcrlf & "Going to delete it.")

       fs.deleteFile pathToFile

    ' }
    else ' {

       msgBox(pathToFile & " does not exists." & vbcrlf & "Going to create it.")

       dim ts as scripting.textStream
       set ts = fs.createTextFile(pathToFile)

       ts.writeLine("First line" )
       ts.writeLine("Second line")
       ts.writeLine("Third line" )

       ts.close

    end if ' }

end sub ' }
