option explicit

'
'  Declare the name of a message …
'
public event message(text as string)

public sub sendMessage(text as string) ' {
  '
  ' Send the specified message text
  '
    raiseEvent message(text)
end sub ' }
