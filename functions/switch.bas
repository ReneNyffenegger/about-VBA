option explicit

sub main()

    dim num as long
    num = 4

    debug.print (switch (        _
        num = 1, "one"      ,    _
        num = 2, "two"      ,    _
        num = 3, "three"    ,    _
        num = 4, "four"     ,    _
        num = 5, "five"     ,    _
        num = 6, "six"      ,    _
        num = 7, "seven"    ,    _
        num = 8, "eight"    ,    _
        num = 9, "nine"     ,    _
        num > 9, "too large")    _
    )

end sub
