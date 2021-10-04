sub addDateIntervals()

    debug.print "Date in a second:       " & dateAdd("s"   , 1, now)
    debug.print "Date in a minute:       " & dateAdd("n"   , 1, now)
    debug.print "Date in a hour:         " & dateAdd("h"   , 1, now)
    debug.print "Date in a week:         " & dateAdd("ww"  , 1, now)
    debug.print "Date in a weekday:      " & dateAdd("w"   , 1, now)
    debug.print "Date in a day:          " & dateAdd("d"   , 1, now)
    debug.print "Date in a month:        " & dateAdd("m"   , 1, now)
    debug.print "Date in a quarter:      " & dateAdd("q"   , 1, now)
    debug.print "Date in a year:         " & dateAdd("yyyy", 1, now)

'   ?                                        dateAdd("y"   , 1, now)

end sub
