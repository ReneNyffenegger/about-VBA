option explicit

sub main ' {

  '
  ' Create two instances that will send messages
  ' messages:

    dim snd_1 as new sender
    dim snd_2 as new sender

  '
  ' Create three receivers
  '
    dim rcv_1 as new receiver
    dim rcv_2 as new receiver
    dim rcv_3 as new receiver

  '
  ' Assign the senders to the receivers:
  '
    rcv_1.init snd_1, "foo"
    rcv_2.init snd_2, "bar"
    rcv_3.init snd_1, "baz"

  '
  ' Send the messages
  '
    snd_1.sendMessage "Message one"
    snd_2.sendMessage "Message two"

end sub ' }
