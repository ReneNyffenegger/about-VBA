option explicit

defLng  i
defInt  j-k
defStr  s

sub main() ' {

  '
  ' Variables that are local to procedures still need to be declared
  '
    dim iNum as long

    for iNum = 1 to 10 ' {
      
        if iNum = 5 then ' {
           S1(iNum)
           S2(iNum)
           S3(iNum)
        end if ' }

    next iNum ' }

end sub ' }

sub S1(iParam) ' {
    debug.print("type of iParam = " & typeName(iParam))
end sub ' }

sub S2(jParam) ' {
    debug.print("type of jParam = " & typeName(jParam))
end sub ' }

sub S3(sParam) ' {
    debug.print("type of sParam = " & typeName(sParam))
end sub ' }

' Output:
'
'   type of iParam = Long
'   type of jParam = Integer
'   type of sParam = String
