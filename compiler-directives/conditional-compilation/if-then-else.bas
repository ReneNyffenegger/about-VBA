option explicit

sub main()

   #if VBA7 then
       debug.print "VBA7 is defined"
   #else
       debug.print "VBA7 is not defined"
   #end if

   #if     win64 then
           debug.print "win64 is defined"
   #elseif win32 then
           debug.print "win32 is defined"
   #else
           debug.print "neither win64 nor win32 is defined"
   #end if

end sub
