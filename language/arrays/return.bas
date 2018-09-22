option explicit

function fibonacci(n as integer) as long() ' {

    dim f() as long
    dim i   as integer

    redim f(1 to n)

    f(1) = 1
    f(2) = 1

    for i = 3 to n ' {
        f(i) = f(i-2) + f(i-1)
    next i ' }

    fibonacci = f

end function ' }

sub main() ' {

   dim ary() as long
   dim i as integer

   ary = fibonacci(10)

   for i = lBound(ary) to uBound(ary) ' {

       debug.print i & ": " & ary(i)

   next i ' }

end sub ' }
