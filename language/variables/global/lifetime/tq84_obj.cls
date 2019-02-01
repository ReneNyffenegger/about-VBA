option explicit

public name_ as string

private sub class_initialize()
    debug.print"tq84_obj initialized"
end sub

private sub class_terminate()
    debug.print "tq84_obj finalized"
end sub

sub sayHello() ' {
    debug.print"Hello world, my name is " & name_
end sub ' }
