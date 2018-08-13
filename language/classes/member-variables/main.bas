option explicit

sub main()

    dim obj as new tq84

    obj.bar = "hello world"

  '
  ' The member baz is declared private, it
  ' cannot be assigned outside of the class:
  '
  ' obj.baz = 99

end sub
