option explicit

sub main() ' {

    dim doc as new MSXML2.DOMDocument

    doc.loadXML("<rootElem><closingTagIsMissing>Not a valid XML document!</rootElem>")

    if doc.parseError.errorCode <> 0 then ' {

       debug.print("Error: " & doc.parseError.reason)
     '
     ' Error: End tag 'rootElem' does not match the start tag 'closingTagIsMissing'.
     '

    end if ' }

end sub ' }
