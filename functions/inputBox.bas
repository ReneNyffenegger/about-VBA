option explicit

sub main() ' {

    dim val as string
    val = inputBox("Enter a value")
    debug.print("You entered: " & val)

    val = inputBox("Enter a value", "Title")
    debug.print("You entered: " & val)

    val = inputBox("Enter a value", "Title", "Default value")
    debug.print("You entered: " & val)

  '
  ' Indicate x and y coordinates
  '
    val = inputBox("Enter a value", "Title", "Default value", 100, 50)
    debug.print("You entered: " & val)

end sub ' }
