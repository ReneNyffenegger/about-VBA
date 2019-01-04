option explicit

sub main() ' {

   dim doc as new MSXML2.DOMDocument

   doc.loadXML(                                                                                     _
     "<items>"                                                                                    & _
     "  <item id='1000'><name val='ABC'/><name val='DEF'/><name val='GHI'/><foo>xxx</foo></item>" & _
     "  <item id='1001'><name val='JKL'/><name val='MNO'/><name val='PQR'/><bar>yyy</bar></item>" & _
     "  <item id='1002'><name val='STU'/><name val='VWX'/><name val='YZ.'/><baz>zzz</baz></item>" & _
    "</items>")

    dim item as msxml2.IXMLDOMElement
    set item = doc.selectSingleNode("//item[@id='1002']")

    dim names as msxml2.IXMLDOMSelection

  '
  ' Selecting the nodes with "//name" gets ALL nodes.
  ' See https://stackoverflow.com/q/54036433/180275
  '
  ' set names = item.selectNodes("//name")
  '
  '
  ' Instead, use a relative path:
  '
    set names = item.selectNodes("./name")

    dim name as msxml2.IXMLDOMElement
    for each name in names
        debug.print(name.getAttribute("val"))
    next name

end sub ' }
