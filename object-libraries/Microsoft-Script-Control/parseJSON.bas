'
'   https://stackoverflow.com/a/37711113  was quite helpful.
'
option explicit

sub main() ' {

    dim jsonText as string

    jsonText =                                        _
    "{"                                        & vbcrlf & _
    "  ""num"" :  42,"                         & vbcrlf & _
    "  ""lst"" : [""foo"", ""bar"", ""baz""]," & vbcrlf & _
    "  ""obj"" : {""hello"": ""World""}"       & vbcrlf & _
    "}"

    debug.print(jsonText)

    dim scrContr as new MSScriptControl.ScriptControl
    scrContr.language = "JScript"

    dim parsed as object ' MSScriptControl.JScriptTypeInfo ?

    set parsed = scrContr.eval("(" & jsonText & ")")

    debug.print("num       = " & parsed.num      )
    debug.print("obj.hello = " & parsed.obj.hello)

  '
  ' Apparently, Arrays are a bit more complicated to access
  '
    dim lenLst as long
    lenLst = callByName(parsed.lst, "length", vbGet)

    dim i as long
    for i = 0 to lenLst-1
        debug.print("lst(" & i & ")    = " & callByName(parsed.lst, i, vbGet))
    next i

end sub ' }
