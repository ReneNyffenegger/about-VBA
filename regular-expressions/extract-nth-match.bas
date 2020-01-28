option explicit

sub main() ' {

    dim txt as string
    txt = "ninety-nine: 99, fourty-two: 42, eleven: 11"

    debug.print extractNthNumber(txt, 1) ' 99
    debug.print extractNthNumber(txt, 2) ' 42
    debug.print extractNthNumber(txt, 3) ' 11

end sub ' }

private function extractNthNumber(txt as string, n as integer) as integer ' {

    dim re as new regExp
    re.pattern   ="\d+"
    re.global    = true
  ' re.multiLine = true

    dim mtcColl as matchCollection

    set mtcColl = re.execute(txt)

    dim mtc as match
    set mtc = mtcColl(n - 1)

    extractNthNumber = mtc.value

end function ' }
