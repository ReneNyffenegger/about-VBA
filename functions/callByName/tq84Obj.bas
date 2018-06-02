'
' Put this code in a class module and rename it to tq84Obj
'
option explicit

public property get funcOne() as string
  funcOne = "this is the first function"
end property

public property get funcTwo() as string
  funcTwo = "this is the second function"
end property

public sub subOne()
  debug.print "first sub was called"
End Sub

public sub subTwo()
  debug.print "second sub was called"
end sub
