option explicit

sub main() ' {

    dim str as string
    str = "foo bar baz"
    s str
    debug.print "After calling s, str = " & str ' changed

end sub ' }

sub s(param as string) ' {

    debug.print "s recevied param = " & param ' foo bar baz
    param = "changed"

end sub ' }
