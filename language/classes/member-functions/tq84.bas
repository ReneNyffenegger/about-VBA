option explicit

private value as long

public sub store_value(value_ as long)

  ' Apprently, the »me« keyword cannot be used
  ' in member subs and member functions …

    value = value_
end sub

public function retrieve_value as long
    retrieve_value = value
end function
