option explicit

private sub class_initialize() ' No parameters possible!
    msgBox "Constructor of class was called"
end sub

private sub class_terminate()
    msgBox "Destructor of class was called"
end sub
