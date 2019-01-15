option explicit

sub main()

  dim dt_1 as date
  dim dt_2 as date
  dim dt_3 as date
  dim dt_4 as date

  dt_1  = #08/20/2010#
  dt_2  = #2010-08-28#

  dt_3  = #2010-08-28 01:02:03#
  dt_4  = #08/28/2010 11:22:33 AM#

  debug.print("There are " & dateDiff("d", dt_1, dt_2) & " days between "    & dt_1 & " and " & dt_2)
' There are 8 days between 20.08.2010 and 28.08.2010

  debug.print("There are " & dateDiff("n", dt_3, dt_4) & " minutes between " & dt_3 & " and " & dt_4)
' There are 620 minutes between 28.08.2010 01:02:03 and 28.08.2010 11:22:33

end sub
