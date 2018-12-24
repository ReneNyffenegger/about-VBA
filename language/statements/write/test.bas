option explicit

sub main() ' {

    dim fileName as string
    fileName = environ$("temp") & "\VBA-print-test.txt"

    dim f as integer
    f = freeFile()

    open fileName for output as #f

    write# f, "Line one"
    write# f, "Line two"
    write# f, "Line three"

    close f

    debug.print("Check " & fileName)

end sub ' }
