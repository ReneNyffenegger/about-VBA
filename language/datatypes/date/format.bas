option explicit

sub main() ' {

    debug.print format(now(), "yyyy-mm-dd"         ) ' 2018-08-27
    debug.print format(now(), "yyyy-mm-dd hh:nn:ss") ' 2018-08-27 10:15:47

    debug.print format(now(), "general date"       ) ' 27.08.2018 10:15:47

    debug.print format(now(), "short date"         ) ' 27.08.2018
    debug.print format(now(), "medium date"        ) ' 27. Aug. 18
    debug.print format(now(), "long date"          ) ' Montag, 27. August 2018

    debug.print format(now(), "short time"         ) ' 10:15
    debug.print format(now(), "medium time"        ) ' 10:15
    debug.print format(now(), "long time"          ) ' 10:15:47

end sub ' }
