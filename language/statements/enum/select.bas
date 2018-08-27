' option explicit

enum META ' {
    FOO
    BAR
    BAZ
end enum ' }

sub main() ' {

    s1 BAR
    s1 2
    s1 5

end sub ' }

sub s1(m as META) ' {

    debug.print "m = " & m

    select case m
       case FOO : debug.print "foo"
       case BAR : debug.print "bar"
       case BAZ : debug.print "baz"
       case else: debug.print "?"
    end select

end sub ' }
'
' Output is:
'
' m = 1
' bar
' m = 2
' baz
' m = 5
' ?
