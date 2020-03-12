option explicit


sub main() ' {

    dim lng as long
    lng = 7

    debug.print( isEmpty  (lng) )  ' False
    debug.print( isArray  (lng) )  ' False
    debug.print( isNumeric(lng) )  ' True
    debug.print( isMissing(lng) )  ' False
    debug.print( isNull   (lng) )  ' False
    debug.print( isObject (lng) )  ' False

    dim var as variant

    debug.print( isArray  (var) )  ' False
    debug.print( isEmpty  (var) )  ' True
    debug.print( isNumeric(var) )  ' True    ' True although not assigned a value
    debug.print( isMissing(var) )  ' False
    debug.print( isNull   (var) )  ' False
    debug.print( isObject (var) )  ' False

    var = "  7.3   "

    debug.print( isArray  (var) )  ' False
    debug.print( isEmpty  (var) )  ' False   ' Because var is now assigned a value
    debug.print( isNumeric(var) )  ' True
    debug.print( isMissing(var) )  ' False
    debug.print( isNull   (var) )  ' False
    debug.print( isObject (var) )  ' False

    var = "  7.3 liters  "

    debug.print( isArray  (var) )  ' False
    debug.print( isEmpty  (var) )  ' False
    debug.print( isNumeric(var) )  ' False   ' spaces are ok, but not letters.
    debug.print( isMissing(var) )  ' False
    debug.print( isNull   (var) )  ' False
    debug.print( isObject (var) )  ' False

    var = null

    debug.print( isArray  (var) )  ' False
    debug.print( isEmpty  (var) )  ' False
    debug.print( isNumeric(var) )  ' False
    debug.print( isMissing(var) )  ' False
    debug.print( isNull   (var) )  ' True    ' Because var is assigned null
    debug.print( isObject (var) )  ' False

    dim obj as excel.worksheet

    debug.print( isArray  (obj) )  ' False
    debug.print( isEmpty  (obj) )  ' False
    debug.print( isNumeric(obj) )  ' False
    debug.print( isMissing(obj) )  ' False
    debug.print( isNull   (obj) )  ' False
    debug.print( isObject (obj) )  ' True

    set obj = activeSheet          '           Assigning a an object does not change returned values

    debug.print( isArray  (obj) )  ' False
    debug.print( isEmpty  (obj) )  ' False
    debug.print( isNumeric(obj) )  ' False
    debug.print( isMissing(obj) )  ' False
    debug.print( isNull   (obj) )  ' False
    debug.print( isObject (obj) )  ' True

end sub ' }
