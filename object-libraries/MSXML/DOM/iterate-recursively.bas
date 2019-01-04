option explicit

sub main() ' {

   dim doc as new MSXML2.DOMDocument

   doc.LoadXML ("<elemRoot>"                                  & _
                "  <child id='one' flagged='yes'>foo</child>" & _
                "  <!-- comment -->"                          & _
                "  <child id='two'>"                          & _
                "      pre bar"                               & _
                "     <grand-child>ABC</grand-child>"         & _
                "      post par"                              & _
                "  </child>"                                  & _
                "  <child id='three'>baz</child>"             & _
                "</elemRoot>")

   call iterateRecursively(doc.childNodes(0), 0)

end sub ' }

sub iterateRecursively(byRef node as variant, byVal level as long) ' {

    dim txt as string

    if node.nodeType = MSXML2.NODE_ELEMENT then
         txt = node.nodeName
    else
         txt = node.Text
    end if

    debug.print (space$(level * 4) & node.nodeTypeString & ": " & txt)

    dim childNode as variant
    dim i as long


    if node.nodeType = MSXML2.NODE_ELEMENT then

       dim attr as MSXML2.IXMLDOMAttribute
       for each attr in node.attributes
           debug.print (space$(level * 4 + 2) & attr.Name & " = " & attr.Value)
       next attr

       for each childNode in node.childNodes
           call iterateRecursively(childNode, level + 1)
       next childNode

    end if

end sub ' }
