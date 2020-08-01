option explicit

sub main() ' {

    dim v1, v2 as long

    debug.print("typeName(v1) = " & typename(v1))
    debug.print("typeName(v2) = " & typename(v2))
    '
    '  typeName(v1) = Empty
    '  typeName(v2) = Long

    v1 = 42
    v2 = 99

    debug.print("typeName(v1) = " & typename(v1))
    debug.print("typeName(v2) = " & typename(v2))
    '
    '  typeName(v1) = Integer
    '  typeName(v2) = Long

    v1 = "v1 is a actually a variant"
 '  v2 = "v2 cannot be assigned a string"

    debug.print("typeName(v1) = " & typename(v1))
    debug.print("typeName(v2) = " & typename(v2))
    '
    '  typeName(v1) = String
    '  typeName(v2) = Long

end sub ' }
