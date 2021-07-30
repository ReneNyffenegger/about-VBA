option explicit

sub main() ' {

     debug.print "A"

#if 0 then

     debug.print "B"

  #if 0 then
     debug.print "C"
  #else
     debug.print "D"
  #end if

     debug.print "E"

#else

  #if 1 then
     debug.print "F"
  #else
     debug.print "G"
  #end if

     debug.print "H"

#end if

     debug.print "I"

end sub ' }
'
' A
' F
' H
' I
