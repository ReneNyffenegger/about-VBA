option explicit

sub main() ' {

    dim arrayOfStrings() as string

    dim i as integer

    arrayOfStrings = split("foo/bar/baz", "/")

    for i = lBound(arrayOfStrings) to uBound(arrayOfStrings)

        debug.print arrayOfStrings(i)

    next i

end sub ' }
