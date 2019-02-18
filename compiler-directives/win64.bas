option explicit

sub main() ' {

#if win64 then
    dim ptr as longLong
#else
    dim ptr as long
#endif

    debug.print ("size of pointer = " & len(ptr))

end sub ' }
