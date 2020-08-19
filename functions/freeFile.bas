option explicit

sub main() ' {

    dim fOne, fTwo, fThree as integer

    fOne   = openFile(environ("TEMP") & "\one.txt"  )
    fTwo   = openFile(environ("TEMP") & "\two.txt"  )
    fThree = openFile(environ("TEMP") & "\three.txt")
    
    print# fOne  , "This is file one"
    print# fTwo  , "This is file two"
    print# fThree, "This is file three"
    
    close# fOne
    close# fTwo
    close# fThree

end sub ' }

function openFile(fileName as string) ' {
  '
  ' Open a file, use freeFile to determine next
  ' file number
  '
  
    dim fn as integer
    
    fn = freeFile()
    open fileName for output as #fn
    
    openFile = fn

end function ' }
