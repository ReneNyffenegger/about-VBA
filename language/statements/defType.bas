option explicit

defLng  i
defInt  j-k
defStr  s

sub main() ' {

  '
  ' Variables that are local to procedures still need to be declared with a dim statement
  '
    dim iVar
    dim jVar
    dim sVar

    dim i : for i = 1 to 10 ' {
      
        if i = 5 then ' {
           S1(i)
           S2(i)
           S3(i)

           debug.print("type of iVar = " & typeName(iVar))
           debug.print("type of jVar = " & typeName(jVar))
           debug.print("type of sVar = " & typeName(sVar))

        end if ' }

    next i ' }

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
'   type of iVar = Long
'   type of jVar = Integer
'   type of sVar = String
