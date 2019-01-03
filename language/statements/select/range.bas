option explicit

sub main() ' {

  debug.print(gaugeNumber(   42))
  debug.print(gaugeNumber(    7))
  debug.print(gaugeNumber(  800))
  debug.print(gaugeNumber(- 111))

end sub ' }

private function gaugeNumber(i as long) as string ' {

  select case i
      case   1 to   9: gaugeNumber = i & " is a small number"
      case  10 to  99: gaugeNumber = i & " is a medium number"
      case 100 to 999: gaugeNumber = i & " is a large number"
      case else     :  gaugeNumber = "I won't say something about " & i
  end select

end function
