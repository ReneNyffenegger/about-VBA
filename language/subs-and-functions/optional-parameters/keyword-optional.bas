' runVBAFilesInOffice.vbs -word %CD%\keyword-optional -c main

sub main() 

  call xyz("one", "two", "three") 
  call xyz("apple", "bananana") 
  call xyz("Berlin", param_3 := "Paris") 

  activeDocument.saved = true 

 end sub 
 

 private sub xyz (                         _ 
                param_1 as string        , _ 
       optional param_2 as string = "foo", _ 
       optional param_3 as string = "bar"  _ 
 ) 

     selection.typeText param_1 & " - " & param_2 & " - " & param_3 & chr(13) 

 end sub 
