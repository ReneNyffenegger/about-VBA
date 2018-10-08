option explicit

sub main() ' {
    debug.print sum              (1,2,3,4,5    )
    debug.print sumViaAnotherFunc(1,2,3,4,5,6,7)
end sub ' }

function sumViaAnotherFunc(paramArray v()) as long ' {
  '
  ' Calling sum().
  ' Since sum() takes a paramArray, it will receive an array with one element
  ' which happens to be the array passed to sumViaAnotherFunc.
  '
    sumViaAnotherFunc = sum(v)
end function ' }

function sum(paramArray v() as variant) as long ' {

    dim i as long

    dim v_ as variant

  '
  ' We need to check if sum was called directly or indirectly.
  ' If called directly, the first element of v is not array,
  ' otherwise it is. Depending on our check, we assign
  ' v or v(0) to v_:
  '
    if isArray(v(0)) then v_ = v(0) _
    else                  v_ = v

  '
  ' Calculating the sum for v.
  '
    sum = 0
    for i = lBound(v_) to uBound(v_)
        sum = sum+v_(i)
    next i

end function ' }
