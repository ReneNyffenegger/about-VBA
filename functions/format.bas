option explicit

sub alignDecimalSeperator(num as variant)

  '
  ' In order to align a number on the decimal seperator, two
  ' format expressions are needed. The first one (##0.00) is
  ' to create a string, the second one (@@@@@@) is to right
  ' align the string:
  '
  
    debug.print format( format(num, "##0.00") , "@@@@@@" )

end sub


sub main()
  
    alignDecimalSeperator( 42     )
    alignDecimalSeperator( 12.34  )
    alignDecimalSeperator(765.4321)

end sub
