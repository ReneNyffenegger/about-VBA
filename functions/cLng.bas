option explicit

sub main()

  ' Spaces around the number are possible
    debug.print cLng("  123  "  )
    
  ' Not possible, type missmatch error
  ' debug.print cLng("  456 xyz")
  
  ' With the &h prefix, the following text is interpreted as
  ' a number's hexadecimal representation.
  ' The following prints 42.
    debug.print cLng("&h2a"     )
    
  ' 17.3 is rounded down to 17
    debug.print cLng("17.3")
    
  ' 17.7 is rounded up to 18
    debug.print cLng("17.7")

end sub
