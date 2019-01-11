option explicit

sub main() ' {

    dim dt_

    dt_ = UTCStringToDate("2018-02-27T15:16:17Z")

    debug.print(format$(dt_, "dd.mm.yyyy - hh:nn:ss"))

end sub ' }

function UTCStringToDate(utc as string) as date ' {

    UTCStringToDate = dateSerial(           _
                         mid(utc, 1, 4),    _
                         mid(utc, 6, 2),    _
                         mid(utc, 9, 2)) +  _
                      timeSerial( _
                         mid(utc,12, 2),    _
                         mid(utc,15, 2),    _
                         mid(utc,18, 2))


end function ' }
