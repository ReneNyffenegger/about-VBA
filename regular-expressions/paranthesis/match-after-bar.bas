option explicit

sub main() ' {

    debug.print numberAfterBar("foo 42 bar 99 baz 10") -- 99
    debug.print numberAfterBar("xxx 42 yyy 99 zzz 10") -- n/a

end sub ' }

function numberAfterBar(txt as string) as string ' {

    dim re as new regExp
    re.pattern = "bar +(\d+)"

    dim mtcColl as matchCollection
    set mtcColl = re.execute(txt)

    if mtcColl.count = 0 then
       numberAfterBar = "n/a"
       exit function
    end if

    dim mtc as match
    set mtc = mtcColl(0)

    numberAfterBar = mtc.subMatches(0)

end function ' }
