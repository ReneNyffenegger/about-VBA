option explicit

sub main() ' {

    dim jsFunc as string

    jsFunc =                                          _
    "function tq84(a, b) {"                    & vbcrlf & _
    "  return a + b;"                          & vbcrlf & _
    "}"

  ' debug.print(jsFunc)

    dim scrContr as new MSScriptControl.ScriptControl
    scrContr.language = "JScript"

    scrContr.addCode(jsFunc)

    dim res as double
    res = scrContr.run("tq84", 18, 24)

    debug.print("res = " & res)

end sub ' }
