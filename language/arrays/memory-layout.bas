option explicit

private declare sub RtlMoveMemory lib "kernel32" alias "RtlMoveMemory" ( _
         dest   as any , _
         src    as any , _
   byVal size   as long  )

private declare sub RtlFillMemory lib "kernel32" alias "RtlFillMemory" ( _
         dest   as any , _
   byVal size   as long, _
   byVal fill   as byte  )

sub main() ' {

    dim bytes(0 to 11) as byte

    RtlFillMemory bytes(0), 12, 0

  '
  ' First num (bytes 0 .. 3)
  '
    bytes( 0) =  42

  '
  ' Second num (bytes 4 .. 7)
  '
  '   4*256 + 176 = 1200
  '
    bytes( 4) = 176
    bytes( 5) =   4

  '
  ' Third num (bytes 8 .. 11)
  '
    bytes( 8) = 255
    bytes( 9) = 255
    bytes(10) = 255
    bytes(11) = 255

    dim longs(0 to 2) as long

    RtlMoveMemory longs(0), bytes(0), 12

    debug.print "longs(0) = " & longs(0)
  '
  ' longs(0) = 42
  '

    debug.print "longs(1) = " & longs(1)
  '
  ' longs(1) = 1200
  '

    debug.print "longs(2) = " & longs(2)
  '
  ' longs(2) = -1
  '

end sub ' }
