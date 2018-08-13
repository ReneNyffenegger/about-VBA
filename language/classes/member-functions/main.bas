option explicit

sub main()

    dim obj as new tq84

    obj.store_value(42)

    msgBox "Stored value is " & obj.retrieve_value

end sub
