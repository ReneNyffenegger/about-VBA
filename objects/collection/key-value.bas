option explicit

dim c as new collection

sub main() ' {

    addItems
    printKeyVal "car"
    printKeyVal "city"
    printKeyVal "fruit"

end sub ' }

sub addItems() ' {

    c.add "Apple"    , "fruit"
    c.add "Mercedes" , "car"
    c.add "New York" , "city"

end sub ' }

sub printKeyVal(key_ as string) ' {
    debug.print key_ & " = " & c(key_)
end sub ' }
