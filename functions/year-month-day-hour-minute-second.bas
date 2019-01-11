option explicit

sub main() ' {

    dim dt as date
    dt = now()

    debug.print("Year:   " & year  (dt))
    debug.print("Month:  " & month (dt))
    debug.print("Day;    " & day   (dt))

    debug.print("Hour:   " & hour  (dt))
    debug.print("Minute: " & minute(dt))
    debug.print("Second: " & second(dt))

end sub ' }
