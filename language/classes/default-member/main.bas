option explicit

sub main() ' {

    dim p as new person

    p.name = "Fred"
    p.age  =    42

    debug.print p.name & "'s age is " & p.age

  '
  ' Use power of default property: name needs
  ' not be specified.
  '
    debug.print p & "'s age is " & p.age

  '
  ' Change name, again using default property
  '
    p = "Charly"
    debug.print p & "'s age is " & p.age

end sub ' }
