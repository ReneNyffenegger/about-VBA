option explicit

sub main() ' {

    dim txt as string

    txt = "one xxx three"
    mid(txt, 5, 3) = "two"
    debug.print txt           ' one two three

    txt = "ONE XY THREE"
    mid(txt, 5, 2) = "TWO"
    debug.print txt           ' ONE TW THREE

end sub ' }
