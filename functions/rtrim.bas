sub main()

    dim txt as string
    txt = "     foo bar baz          "
    
    debug.print "txt: >" & txt & "<"
  
  '
  ' Traim trailing spaces on the right side:
  '  
    txt = rTrim(txt)
    
    debug.print "txt: >" & txt & "<"

end sub