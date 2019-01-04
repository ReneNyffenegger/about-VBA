option explicit

sub main() ' {

   dim doc as new MSXML2.DOMDocument

   doc.loadXML(                             _
     "<items>"                            & _
     "  <fruit id='1000'>Apple</fruit>"   & _
     "  <fruit id='1002'>Cherry</fruit>"  & _
     "  <city  id='2000'>Paris</city>"    & _
     "  <fruit id='1001'>Banana</fruit>"  & _
     "  <city  id='2001'>Berlin</city>"   & _
     "  <city  id='2002'>Zürich</city>"   & _
     "</items>")


    dim fruits as msxml2.IXMLDOMSelection
    set fruits = doc.selectNodes("//fruit")

    dim fruit as msxml2.IXMLDOMElement
    for each fruit in fruits
        debug.print("Fruit with id " & fruit.getAttribute("id") & " is " & fruit.text)
    next fruit

end sub ' }
