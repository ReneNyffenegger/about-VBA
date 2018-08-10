option explicit

sub determineType(obj as variant)

    if obj is nothing then
       debug.print "nothing"

     ' If the object is »nothing«, it makes no sense to continue
     ' with typeOf … is checks: nothing-objects cause error message:
     '   Object variable or With block variable not set.
     '
     ' Thus, we exit the sub:
       exit sub
    end if

    if typeOf obj is object then
       debug.print "object"
    end if

    if typeOf obj is classOne then
       debug.print "classOne"
    end if

    if typeOf obj is classTwo then
       debug.print "classTwo"
    end if

end sub

sub main()

    dim obj_1 as     classOne
    dim obj_2 as new classTwo

    determineType obj_1
    determineType obj_2

end sub
