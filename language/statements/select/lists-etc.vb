option explicit

sub main()

    X "Hello world"
    X "FOO"
    X "foo"
    X "alpha"
    X "Beta"
    X "One"
    X "one"

end sub

sub X(val as string)

    dim res as string
    select case val

       case "foo", "bar", "ba"
            res = "metasyntactic variable"

       case "A" to "G", "X" to "Z", lcase(val) = "one"
          '
          ' Note lcase(val) = "one" probably does not what was
          ' intended.
          '
            res = "between A and G, between X and Z or true"

       case else
            res = "> G"

     end select

     debug.print(val & ": " & res)

end sub
'
' Hello world: > G
' FOO: between A and G, between X and Z or case insenstive 'one'
' foo: metasyntactic variable
' alpha: > G
' Beta: between A and G, between X and Z or case insenstive 'one'
' One: > G
' one: > G
