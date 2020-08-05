option explicit

sub main() ' {

    dim outer as long

    for outer = 1 to 5 ' {

        dim inner as long

        inner = inner + 1
        debug.print "outer = " & outer & ", inner = " & inner

    next outer ' }

end sub ' }
'
' Output is:
'
'   outer = 1, inner = 1
'   outer = 2, inner = 2
'   outer = 3, inner = 3
'   outer = 4, inner = 4
'   outer = 5, inner = 5
