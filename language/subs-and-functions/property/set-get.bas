option explicit

dim someStrings() as string

sub main() ' {
    redim someStrings(0 to 0)

    val(4) = "four"
    val(2) = "two"
    val(1) = "one"
    val(6) = "six"

    dim i as long
    for i = 0 to 10
        debug.print i & ": " & val(i)
    next i

end sub ' }

property let val(byRef item as long, byVal value as string) ' {

    if item > uBound(someStrings) then
       reDim preserve someStrings(0 to item)
    end if

    someStrings(item) = value

end property ' }

property get val(byRef item as long) as string ' {
    if item > uBound(someStrings) then
       val = "n/a"
       exit property
    end if

    val = someStrings(item)
end property ' }
