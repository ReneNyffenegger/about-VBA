option explicit

sub main() ' {

    dim xyz(0 to 3, _
            0 to 4, _
            0 to 5) as long

    dim x as long, y as long, z as long

    for x = lbound(xyz, 1) to ubound(xyz, 1)
    for y = lbound(xyz, 2) to ubound(xyz, 2)
    for z = lbound(xyz, 3) to ubound(xyz, 3)

           xyz(x, y , z) = x*y + 2*z + y
           debug.print("xyz(" & x & ", " & y & ", z) = " & xyz(x, y, z))

    next z
    next y
    next x

end sub ' }
