option explicit

type tp
    foo as long
    bar as date
end type

sub printTypeName(v as variant)
   debug.print typeName(v)
end sub

sub main()

    dim tp_ as tp
    dim obj as new tq84
    dim var as variant

    printTypeName 42          ' Integer
    printTypeName true        ' Boolean
    printTypeName nothing     ' Nothing
    printTypeName empty       ' Empty
    printTypeName null        ' Null
    printTypeName obj         ' tq84
    printtypeName obj.aMember ' Double

    printTypeName var         ' Empty
    var = #08/28/2010#
    printTypeName var         ' Date

  ' In Excel, it's possible to use typeName to
  ' determine what the application.selection property
  ' actually referrs to. Chances are that it is a range.
    printTypeName selection   ' Range

  ' printTypeName missing     ' error: Variable not defined …

  ' printTypeName tp_  ' Apparently, this is not possible
  '                    ' Error message is: Only user-defined types defined in public object modules
  '                                        can be coerced to or from a variant or passed to late-bound functions

end sub
