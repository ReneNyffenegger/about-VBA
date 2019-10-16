option explicit

sub s0() ' {
    debug.print("s0")
end sub ' }

sub s1(p1 as long) ' {
    debug.print("s1, p1 = " & p1)
end sub ' }

sub s2(p1 as long, p2 as long) ' {
    debug.print("s1, p1 = " & p1 & ", p2 = " & p2)
end sub ' }

function f0() as long' {
    debug.print("f0")
end function ' }

function f1(p1 as long) as long ' {
    debug.print("f1, p1 = " & p1)
end function ' }

function f2(p1 as long, p2 as long) as long ' {
    debug.print("f1, p1 = " & p1 & ", p2 = " & p2)
end function ' }

sub main() ' {

         s0
  '      s0()      ' Parantheses not allowed
    call s0
    call s0()

         s1 1
         s1(1)
  ' call s1 1      ' Parantheses required
    call s1(1)

         s2 1, 2
  '      s2(1, 2)  ' Parantheses not allowed
  ' call s2 1, 2   ' Parantheses required
    call s2(1, 2)

  ' -------------------------------------------

         f0
  '      f0()      ' Parantheses not allowed
    call f0
    call f0()

         f1 1
         f1(1)
  ' call f1 1      ' Parantheses required
    call f1(1)

         f2 1, 2
  '      f2(1, 2)  ' Parantheses not allowed
  ' call f2 1, 2   ' Parantheses required
    call f2(1, 2)

  ' -------------------------------------------

    dim r as long
    r =  f0
    r =  f0()      ' Parantheses allowed (as opposed to when not assigning return value to variable)

  ' r =  f1 1      ' Parantheses required
    r =  f1(1)

  ' r =  f2 1, 2   '
    r =  f2(1, 2)  ' Parantheses allowed (as opposed to when not assigning return value to variable)

end sub ' }
