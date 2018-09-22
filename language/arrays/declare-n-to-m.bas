option explicit

sub main() ' {

    dim ary(0 to 2) as integer
    dim i           as integer
    dim sum         as integer

    ary(0) = 42
    ary(1) = 99
    ary(2) = -8

    sum = 0
    for i = 0 to 2
        sum = sum + ary(i)
    next i

    debug.print "sum = " & sum

end sub ' }
