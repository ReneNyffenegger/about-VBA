sub main()

   printNumber  2
   printNumber  1
   printNumber 99

end sub

sub printNumber(nr as integer)

    debug.print choose(nr, "one", "two", "three")

end sub
