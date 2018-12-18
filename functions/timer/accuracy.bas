option explicit

sub main()

  dim t0, t1       as single
  dim cnt          as long
  dim measurements as long
  
  cnt          = 0
  measurements = 0
  
  t0  = timer
  
  do while cnt < 10
  
     t1           = timer
     measurements = measurements + 1
  
     if t1 > t0 then
        debug.print "Diff: " & round(t1-t0, 3)
        t0 = t1
        cnt = cnt + 1
     end if
  
  loop
  
  debug.print "Total measurements: " & measurements

end sub
