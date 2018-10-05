'
'   Inspired by https://usefulgyaan.wordpress.com/2013/06/12/vba-trick-of-the-week-slicing-an-array-without-loop-application-index/
'
option explicit

sub main() ' {

    dim table    (1 to 9, 1 to 9) as string
    dim ary1d                     as variant

  '
  ' Set up a simple table and store it
  ' in a two dimensional array (variable table):
  '
    dim i, j as long
    for i = 1 to 9: for j = 1 to 9
        table(i, j) = i & "/" & j
    next j: next i

  '
  ' Extract a one dimensional array from the two
  ' dimensional array:
  '
    ary1d = worksheetFunction.index(table, 4, 0)
    debug.print join(ary1d, " - ")

end sub ' }
