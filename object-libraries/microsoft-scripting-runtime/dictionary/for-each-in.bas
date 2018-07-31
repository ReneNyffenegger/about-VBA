option explicit

sub main() ' {

    dim capital as new scripting.dictionary

    capital.add "France"     , "Paris"
    capital.add "Germany"    , "Berlin"
    capital.add "Switzerland", "Bern"
    capital.add "Italy"      , "Rome"

    dim country as variant

    for each country in capital
        debug.print country & ": " & capital(country)
    next country

end sub ' }
