option explicit

sub main() ' {

    dim bl as boolean
    dim dt as date
    dim em as variant ' Without assignment, a variant variable has teh value empty.
'   dim er as error
    dim nl as variant
    dim nm as long

    dt = 123456.789
    nl = null
    nm = 42

    debug.print cStr(bl ) ' False
    debug.print cStr(dt ) ' 03/01/2238 18:56:10
    debug.print cStr(em ) '
    debug.print cStr(err) ' 40040
 '  debug.print cStr(nl ) ' runtime error '94': Invalid use of Null
    debug.print cStr(nm ) ' 42

end sub ' }
