option explicit

sub main() ' {

    dim dtStart as date
    dim dtEnd   as date

    dtStart = dateSerial(2018,  8, 18)
    dtEnd   = dateSerial(2018,  8, 28)

    dim daysBetween as long

    daysBetween = dateDiff("d", dtStart, dtEnd)

    debug.print("There are " & daysBetween & " days between " & dtStart & " and " & dtEnd)
  '
  ' There are 10 days between 18.08.2018 and 28.08.2018
  '

end sub ' }
