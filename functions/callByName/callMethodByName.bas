option explicit

sub main()

    dim obj   as tq84Obj
    set obj = new tq84Obj
    
    dim res as string
    
    debug.print CallByName(obj, "funcOne", vbGet)
    debug.print CallByName(obj, "funcTwo", vbGet)

    callByName obj, "subOne", vbMethod
    callByName obj, "subTwo", vbMethod
    
    debug.print res

end sub
