option explicit

sub printHourMinuteSecond() ' {

    dim secsSinceMidnight as single ' timer returns also fractional part.
    dim hour_       as integer
    dim minute_     as integer
    dim second_     as integer
    dim secondPart_ as single

    secsSinceMidnight = timer

    hour_       = int(  secsSinceMidnight        /60/60                )
    minute_     = int( (secsSinceMidnight - hour_*60*60 )        /60   )
    second_     = int( (secsSinceMidnight - hour_*60*60 - minute_*60 ) )
    secondPart_ =       secsSinceMidnight - hour_*60*60 - minute_*60 - second_

    debug.print("Calculated hr:mi:ss is: " & _
       format(hour_      ,  "00"  ) & ":"  & _
       format(minute_    ,  "00"  ) & ":"  & _
       format(second_    ,  "00"  )        & _
       format(secondPart_, ".00") )

end sub ' }
