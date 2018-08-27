option explicit

sub main() ' {

    double_as_date 0.0                                       ' 1899-12-30 00:00:00
    double_as_date 0.5                                       ' 1899-12-30 12:00:00

    double_as_date 365.0                                     ' 1900-12-30 00:00:00

    double_as_date 43339                                     ' 2018-08-27 00:00:00
    double_as_date 43339 + 11/24 + 27/24/60 + 39/24/60/60    ' 2018-08-27 11:27:39

end sub ' }

sub double_as_date(dbl as double) ' {

    dim dt as date
    dt = dbl
    debug.print format(dt, "yyyy-mm-dd hh:nn:ss")

end sub ' }
