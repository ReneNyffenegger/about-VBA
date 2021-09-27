option explicit

sub main() ' {

    dim jsFunc as string

    cells(2,2) = "B2"
    cells(2,3) = "B3"
    cells(3,2) = "C2"
    cells(3,3) = "C3"

    jsFunc = _
      "function tq84(val, rng) {" & vbCrLf & _
      "   return 'val = ' + val+ ', rng.address = ' + rng.address;" & vbCrLf & _
      "}"

  ' debug.print(jsFunc)

    dim scrContr as new MSScriptControl.ScriptControl
    scrContr.language = "JScript"

    scrContr.addCode(jsFunc)

    dim res as string
    res = scrContr.run("tq84", 42, range(cells(2,2), cells(3,3)))

    debug.print("res: " & res)

end sub '
