option explicit

sub main() ' {

   dim doc as new MSXML2.DOMDocument

   doc.loadXML(                             _
     "<items>"                            & _
     "  <fruit id='1000'>Apple</fruit>"   & _
     "  <fruit id='1001'>Banana</fruit>"  & _
     "  <fruit id='1002'>Cherry</fruit>"  & _
     "  <city  id='2000'>Paris</city>"    & _
     "  <city  id='2001'>Berlin</city>"   & _
     "  <city  id='2002'>Zürich</city>"   & _
     "</items>")


    debug.print(doc.selectSingleNode("//*[@id='1001']").text) ' Banana

    debug.print(doc.selectSingleNode("/items/city[2]").text)  ' Zürich

end sub ' }
