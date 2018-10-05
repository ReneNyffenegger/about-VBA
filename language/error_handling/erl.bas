option explicit

sub main() ' {
  1
  2   on error goto err_
  3
  4   debug.print "in main"
  5   debug.print "This sub has line numbers"
  6   debug.print "which make it possible to"
  7   debug.print "report the line when an"
  8   debug.print "error occurs."
  9
 10   dim i as long
 11   i = 42 / 0          ' Division by Zero!
 12
 13   debug.print "This line is not reached"
 14   debug.print "because the assignment to"
 15   debug.print "the varialbe i causes an error."
 16
 17   exit sub
 18
 19   err_:
 20     debug.print "*** Error ***"
 21     debug.print "   " & err.description & " on line " & erl
 22
end sub ' }
