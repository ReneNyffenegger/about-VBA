option explicit

sub main() ' {

    readFile "latin-1.txt"
    readFile "utf-8.txt"
    readFile "utf-16.txt"

end sub ' }

sub readFile(fileName as string) ' {

   const directoryName$ = "C:\Users\r.nyffenegger\personal\VBA\open-charset\"

   dim f as integer
   f = freeFile()

   open directoryName & fileName for input as #f

   dim txt as string
   line input #f, txt

   debug.print ("Line read from " & fileName & " is: " & txt)

   close f

end sub ' }
