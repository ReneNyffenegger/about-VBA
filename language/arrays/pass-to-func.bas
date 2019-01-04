option explicit


sub main() ' {

    dim v1(1 to 3) as double
    dim v2(1 to 3) as double

    v1(1) =  1
    v1(2) =  3
    v1(3) = -5

    v2(1) =  4
    v2(2) = -2
    v2(3) = -1

    debug.print(dot_product_explicit_datatype(v1             , v2               ))
    debug.print(dot_product_variant          (array(1, 3, -5), array(4, -2, -1) ))
    debug.print(dot_product_variant          (array(1, 3, -5), array(4, -2, -1) ))

end sub ' }


function dot_product_explicit_datatype (byRef vec_1() as double, byRef vec_2() as double) as double ' {

    dim i as long

    dot_product_explicit_datatype = 0
    for i = lBound(vec_1) to uBound(vec_1)
        dot_product_explicit_datatype = dot_product_explicit_datatype + vec_1(i) * vec_2(i)
    next i

end function ' }


function dot_product_variant (byRef vec_1 as variant, byRef vec_2 as variant) as double ' {

    dim i as long

    dot_product_variant = 0
    for i = lBound(vec_1) to uBound(vec_1)
        dot_product_variant = dot_product_variant + vec_1(i) * vec_2(i)
    next i

end function ' }
