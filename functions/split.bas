option explicit

sub splitDemo()

    dim text as string
    text = "foo,bar,baz"

    dim ary() as string
    ary = split(text, ",")

  ' Note that the returned array is
  ' zero-based. Therefore, we need to
  ' start iterating with i = 0
    for i = 0 to uBound(ary)
        msgBox(ary(i))
    next i

end sub
