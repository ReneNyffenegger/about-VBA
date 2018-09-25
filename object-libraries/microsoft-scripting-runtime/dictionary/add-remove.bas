option explicit

sub main() ' {

  on error goto err_

    dim dict as new dictionary

    debug.print "adding k1"
    dict("k1") = "v1"

    debug.print "adding k2"
    dict.add "k2", "v2"

    debug.print "adding k1"
    dict.add "k1", "...."

    debug.print "removing k1"
    dict.remove "k1"

    debug.print "removing inexisting-key"
    dict.remove "inexisting-key"

    debug.print "removing k2"
    dict.remove "k2"

    exit sub

  err_:

    if err.number = 457 then
       debug.print "Error: key is already present in dictionary!"
       resume next
    end if

    if err.number = 32811 then
       debug.print "Error: trying to remove key that does not exist!"
       resume next
    end if

    debug.print "unhandled error " & err.number & ": " & err.description

end sub ' }
