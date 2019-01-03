option explicit

sub main() ' {

  debug.print(" 1: " & num_to_text( 1))
  debug.print(" 5: " & num_to_text( 5))
  debug.print("42: " & num_to_text(42))

end sub ' }

private function num_to_text(i as long) as string ' {

  select case i
      case 1   : num_to_text = "one"
      case 2   : num_to_text = "two"
      case 3   : num_to_text = "three"
      case 4   : num_to_text = "four"
      case 5   : num_to_text = "five"
      case 6   : num_to_text = "six"
      case 7   : num_to_text = "seven"
      case 8   : num_to_text = "eight"
      case 9   : num_to_text = "nine"
      case else: num_to_text = "Too big!"
  end select

end function ' }
