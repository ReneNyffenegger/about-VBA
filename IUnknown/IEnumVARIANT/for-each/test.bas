option explicit

sub main() ' {

    dim dbl     as variant
    dim varIter as new variantIterator
    
  ' Show
    for each dbl in varIter.values("one", 2, 3.333)
        debug.print dbl
    next dbl

end sub ' }