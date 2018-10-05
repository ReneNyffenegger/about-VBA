option explicit

sub main() ' {

    dim ary(1 to 3) as string

    ary(1) = "foo"
    ary(2) = "bar"
    ary(3) = "baz"

    dim txt as string
    txt = join(ary, " - ")

    debug.print txt

end sub ' }
