option explicit

sub main()

    dim numbers() as string
    numbers = split("one,two,three,four,five,six,seven,eight,nine,ten", ",")

  '
  ' Find words that contain 've' as substring
  '
    debug.print join(filter(numbers, "ve"), chr(10))

end sub
