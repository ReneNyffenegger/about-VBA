option explicit

sub main() ' {

    dim fileName as string
    fileName = environ$("temp") & "\without-newlines.txt"

    dim f as integer
    f = freeFile()

    open fileName for output as #f

    print# f, "first line: "  ;
    print# f, "one "          ;
    print# f, "two "          ;
    print# f, "three"           ' Note: no semicolon to add new line
    print# f, "second line: " ;
    print# f, "foo "          ;
    print# f, "bar "          ;
    print# f, "baz"             ' write another new line

    close f

    debug.print("Check " & fileName)

end sub ' }
