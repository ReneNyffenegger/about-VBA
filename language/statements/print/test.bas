option explicit

sub main() ' {

    dim fileName as string
    fileName = environ$("temp") & "\VBA-print-test.txt"

    dim f as integer
    f = freeFile()

    open fileName for output as #f

    print# f, "Line one"
    print# f, "Line two"
    print# f, "Line three"

    close f

    debug.print("Check " & fileName)

end sub ' }
