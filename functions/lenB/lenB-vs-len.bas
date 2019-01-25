option explicit

type FOO ' {
     b1  as byte
     b2  as byte
     itg as long
     b3  as byte
end type  ' }

sub main() '

    dim byt as byte     : byt =      42    : debug.print("byte     : " & len(byt) & " " & lenB(byt))
    dim bol as boolean  : bol =    true    : debug.print("boolean  : " & len(bol) & " " & lenB(bol))
    dim cur as currency : cur =      18.05 : debug.print("currency : " & len(cur) & " " & lenB(cur))
    dim dat as date     : dat =    now()   : debug.print("date     : " & len(dat) & " " & lenB(dat))
    dim dbl as double   : dbl =     271.38 : debug.print("double   : " & len(dbl) & " " & lenB(dbl))
    dim itg as integer  : itg =   12345    : debug.print("integer  : " & len(itg) & " " & lenB(itg))
    dim lng as long     : lng =  123456    : debug.print("long     : " & len(lng) & " " & lenB(lng))
  ' dim llg as long long:                  : debug.print("long long: " & len(llg) & " " & lenB(llg))
    dim sng as single   : sng =      18.73 : debug.print("single   : " & len(sng) & " " & lenB(sng))
    dim str as string   : str =    "text"  : debug.print("string   : " & len(str) & " " & lenB(str))
    dim typ as FOO      :                    debug.print("FOO      : " & len(typ) & " " & lenB(typ))
  ' dim obj as object   "                    debug.print("object   : " & len(obj) & " " & lenB(obj))
    dim var as variant  : var =      15    : debug.print("variant  : " & len(var) & " " & lenB(var))

end sub ' }
'
' byte     : 1 1
' boolean  : 2 2
' currency : 8 8
' date     : 8 8
' double   : 8 8
' integer  : 2 2
' long     : 4 4
' single   : 4 4
' string   : 4 8
' FOO      : 7 12
' variant  : 2 4
'
