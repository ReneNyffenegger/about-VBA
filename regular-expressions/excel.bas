option explicit

sub main() ' {

    fillTestData

    range(cells(1,6), cells(4,6)).formulaR1C1 = "=extractNumber(rc[-5])"

end sub ' }

sub fillTestData() ' {

    cells(1, 1) = "foo bar baz"
    cells(2, 1) = "bla 42 more bla"
    cells(3, 1) = "number 37 another number 99 xyz"
    cells(4, 1) = "line one" + vbCrLf + "line two" + vbCrLf + "line three (3)" + vbCrLf + "line four"

end sub ' }

public function extractNumber(cell as range) as string ' {

    dim re as new regExp
    re.pattern   ="\d+"
  ' re.global    = true
  ' re.multiLine = true

    dim mtcColl as matchCollection

    set mtcColl = re.execute(cell)

    dim mtc as match
    set mtc = mtcColl(0)

    extractNumber = mtc.value

end function ' }
