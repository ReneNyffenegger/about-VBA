option explicit

sub main() ' {

    call doBoth("foo bar baz"                                 )
    call doBoth(""                                            )
    call doBoth("one  two     three four   five     six seven")

end sub ' }

sub doBoth(txt as string) ' {

     call matchUpToSixWords(txt)
     call matchAllWords    (txt)

end sub ' }

sub matchUpToSixWords(txt as string) ' {

    debug.print("Get up to 6 words of """ & txt & """")

    dim re as new regExp
    re.pattern = "(\w+)?( +)?" & _
                 "(\w+)?( +)?" & _
                 "(\w+)?( +)?" & _
                 "(\w+)?( +)?" & _
                 "(\w+)?( +)?" & _
                 "(\w+)?"


    dim match as matchCollection
    set match = re.execute(txt)

 '  debug.print(typeName(match)) ' IMatchCollection2

 '  debug.print("Count of matches =    " & match.count)

    dim parantheses as subMatches
    set parantheses = match(0).submatches

 '  debug.print(typeName(parantheses)) ' ISubMatches

 '  debug.print("Count of submatches = " & parantheses.count) ' 11 because there are 11 parantheses-pairs.

    dim i as long
    for i = 0 to 11 step 2

        if not isEmpty(parantheses(i)) then
           debug.print("  " & parantheses(i))
        end if

    next i


end sub ' }

sub matchAllWords(txt as string) ' {

    debug.print("Get all words of """ & txt & """")

    dim re as new regExp

    re.global = true               ' Kind of important...

    re.pattern = "\w+ *"

    dim matches as matchCollection
    set matches = re.execute(txt)

 '  debug.print("matches.count = " & matches.count)

    dim word as variant
    for each word in matches
        debug.print("  " & word)
    next word


end sub ' }
