option explicit


sub main() ' {

    dim obj as new X

    debug.print("--------- start loops ----------")

    loop_1
    loop_2
    loop_3

    debug.print("--------- loops finished -------")

    debug.print("next statement: print y.value")
    obj.doSomething

end sub ' }

sub loop_1() ' {

    debug.print("--------- loop_1 ---------------")

    dim i as long
    for i = 1 to 3
        dim obj as X
        set obj = new X
        debug.print("  i = " & i)
        obj.doSomething
    next i

end sub ' }

sub loop_2() ' {
    debug.print("--------- loop_2 ---------------")

    dim i as long

    dim obj as X
    for i = 1 to 3
        debug.print("  i = " & i)
        set obj = new X
        obj.doSomething
    next i

end sub ' }

sub loop_3() ' {
    debug.print("--------- loop_3 ---------------")

    dim i as long

    for i = 1 to 3
        dim obj as new X
        debug.print("  i = " & i)
        obj.doSomething
    next i

end sub ' }
