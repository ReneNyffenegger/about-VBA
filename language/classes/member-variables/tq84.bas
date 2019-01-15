option explicit

'
' Declare three member variables for the class.
'
' Note: member variables are not declared with dim, rather
' one of the keywords public or private is used.
'
' If they're public, they can be accessed from »outside« the
' class. If private, they can only be accessed within the class.
'

public  foo as double
public  bar as string
private baz as long

private sub class_initialize()
  '
  ' Within a class method, members of the class can
  ' optionally be referred to with »me«. In other
  ' programming languages, that would be self or this.
  '
    me.foo = 1.23456
       baz = 42
end sub

private sub class_terminate()
    msgBox "foo = " & foo & ", bar = " & bar & ", baz = " & baz
end sub
