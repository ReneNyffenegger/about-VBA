option explicit

#const txt = "bar"
#const num =  42

sub testDirectives()

   #if     txt = "foo" then
           debug.print "txt was defined to foo"
   #elseif txt = "bar" then
           debug.print "txt was defined to bar"
   #elseif txt = "baz" then
           debug.print "txt was defined to baz"
   #else
           debug.print "txt is either undefined or not defined to either foo, bar or baz"
   #end if

   #if     num         then
           debug.print "num was defined"

          #if num = 99 then
              debug.print "The value of num = 99"
          #else
              debug.print "The value of num <> 99"
          #end if
   #else
           debug.print "num was not defined"
   #end if

   #if     hello       then
           debug.print "hello was defined"
   #else
           debug.print "hello was not defined"
   #end if

end sub
'
' Output:
'
'   txt was defined to bar
'   num was defined
'   The value of num <> 99
'   hello was not defined
'
