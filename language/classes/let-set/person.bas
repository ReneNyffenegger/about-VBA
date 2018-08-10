'
'      Beware the »Definitions of property procedures for the same property are inconsistent« error.
'
option explicit

private nm as string
private ag as long


property let name (name_ as string)

  '
  ' Apparently, in property let/set functions,
  ' the »me« keyword is not allowed …
  '

    nm = name_
end property

property let age (age_   as long  )
    ag = age_
end property

property get name() as string
    name = nm
end property

property get age() as long
    age = ag
end property

property get text() as string
    text = "My name is " & nm & " and I am " & age & " years old."
end property
