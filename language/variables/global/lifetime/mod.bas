option explicit

global g_obj as tq84_obj
global g_num as double

sub main() ' {
    set g_obj = new tq84_obj
    g_num     = 42
    g_obj.name_ = "main"
end sub ' }

sub finalize() ' {
    debug.print "finalize 1"
    end
    debug.print "finalize 2"
end sub ' }

function causeError(num as double) as double ' {
    causeError = 5 / num
end function ' }
