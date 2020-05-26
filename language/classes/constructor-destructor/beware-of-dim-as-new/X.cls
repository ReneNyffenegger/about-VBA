option explicit

private value as long

private sub class_initialize() ' No parameters possible!
    debug.print "  class_initialize was called"
    value = 42
end sub

private sub class_terminate()
    debug.print "  class_terminate was called"
end sub

public sub doSomething() ' {
    value = value + 1
    debug.print("  value = " & value)
end sub ' }
