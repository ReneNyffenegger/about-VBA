option explicit

sub main()

   dim str as variant
   dim lng as variant

   str = "42"
   lng =   8

   debug.print vbaVarCat(str, lng) ' 842
   debug.print vbaVarAdd(str, lng) '  50

end sub
