public   var_pub  as long           ' Declaring a public variable  (can be used in all modules)
private  var_prv  as date           ' Declaring a private variable (can only be used in THIS module)
const    c_word   = "foo"           ' Declaring a constant. Type of variable is determined implicitely

sub dimTest()

    dim var           as string     ' Declaring a "normal" variable
    dim str_3         as string * 3 ' Fixed width string
    dim ary()         as integer    ' Declaring an array with variable count of elements
    dim ary_5(1 to 5) as integer    ' Declaring a fixed count of elements (here: 5)
    
    call staticVar
    call staticVar

end sub

sub staticVar()

    static var as integer ' Static variables will keep their value accross sub/function calls
                          ' They are initialized with a default value (here: 0)
    
    debug.print "static var = " & var
    var = var + 1

end sub