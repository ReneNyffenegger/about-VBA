option explicit

sub XX()

  dim cbAddr as long
   
  '
  '  addressOf is an operator, not a function.
  '  Therefore, the following is not possible,
  '  it causes a "Compile error, Syntax error"
  '
  '   cbAddr = addressOf callBackSub
  '
  '  However, cbAddr can be assigned the address of
  '  a call-back function or sub with the following
  '  construct:
  
  cbAddr = getAddressOfCallback(addressOf callBackSub)

end sub


function getAddressOfCallback(addr as long) as long
         getAddressOfCallback=addr
end function


sub callBackSub()

    msgBox "xyz"

end sub
