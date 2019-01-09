option explicit

const g_forty_two = 42

private const g_seven = 7
public  const g_five  = 5

'
' It is possible to assign a calculated value to
' a constant:
'
const calculated = g_seven * g_five

'
' However, it is not possible to assign a value to a
' constant as returned value from a function. The following
' statement would cause the error:
'
'      Constant expression required
'
' const anotherCalculated = someNumber()


'
' When declaring constants, their data type can be
' explicitely stated. Without declaring the datatype,
' the data type of the following constant would be integer:
'
const dataTypedConst as long = 10


sub main() ' {

   '
   ' Constants, obviously, cannot be changed.
   ' The following statements would cause the error:
   '
   '    Assignment to constant not permitted
   '
   ' g_forty_two = 45
   ' g_seven = 9
   ' g_five  = 18

   debug.print ("Calculated = " & calculated)

   debug.print(typeName(g_seven       )) ' Integer
   debug.print(typeName(dataTypedConst)) ' Long

end sub ' }

function someNumber() ' {
    someNumber = 99
end function ' }
