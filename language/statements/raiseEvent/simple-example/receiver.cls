option explicit


  '
  ' This class wants to receive messages that
  ' were sent by an instance of the sender class.
  ' The following declaration specifies (a reference)
  ' to the instance which will send the messages:
  '
private withEvents snd as sender

private            id  as string


public sub init(snd_ as sender, id_ as string) ' {
  '
  ' We need to assign the reference to our member
  ' variable snd in order to connect messages
  ' that were sent from the instance to our
  ' member function snd_message (below)

    set snd = snd_
        id  = id_

end sub ' }

private sub snd_message(text as string) ' {
  '
  '  This is the method that receives the events that were
  '  sent with raiseEvent(message).
  '
  '  The method name consists of two parts that are separated by
  '  an underscore. The first part corresponds to the name of the
  '  member variable that was declared with »withEvents« (snd), the
  '  second part is the name of the event that we have an interest in.
  '
     debug.print id & " received " & text
end sub ' }
