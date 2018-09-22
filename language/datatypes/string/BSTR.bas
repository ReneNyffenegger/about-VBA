option explicit

sub main() ' {

    dim str         as string
    dim addr        as longPtr
    dim strLenBytes as long
    dim addrChar    as longPtr
    dim char        as integer

    str = "twenty-one characters"

    addr = strPtr(str)

  '
  ' The length of the string in bytes is stored in the four bytes
  ' prior to the string address:
  '
    strLenBytes = GetMem4_(addr - 4)
    debug.print "strLenBytes = " & strLenBytes

  '
  ' Iterating over each character in the string. Since each character is
  ' two bytes wide, we use "step 2"
  '
    for addrChar = addr to addr + strLenBytes step 2 ' {

      '
      ' Get the numeric value of a character:
      '
        char = GetMem2_(addrChar)

      '
      ' Print the character (using chr$()) and its
      ' numeric value:
      '
        debug.print "chr$(char) = " & chr$(char) & " | " & char

    next addrChar ' }

end sub ' }
