option explicit

sub main()

   dim p1 as new person
   dim p2 as new person

   p1.name = "Fred" : p1.age = 71
   p2.name = "Jane" : p2.age = 68

   msgBox p1.text & chr(13) & p2.text

end sub
