'
'  vim: ft=basic
'
option explicit

implements Iexample

public sub Iexample_aSub()
    msgBox "foo says hello"
end sub

public function Iexample_aFunc(param as double)
    Iexample_aFunc = param*2
end function
