option explicit

sub main() ' {
    debug.print sum(5.1, 2.29, -3.97, 8)
end sub ' }

function sum(paramArray nums()) as double ' {

    dim i as long

    sum = 0

    for i = lBound(nums) to uBound(nums)
        sum = sum + nums(i)
    next i

end function ' }
