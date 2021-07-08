option explicit

implements Iexample

public sub Iexample_aSub()
    msgBox "bar says hello"
end sub

public function Iexample_aFunc(param as double)
    Iexample_aFunc = param*5
end function
