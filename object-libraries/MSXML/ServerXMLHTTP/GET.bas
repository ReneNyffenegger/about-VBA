option explicit

sub main() ' {

   dim httpReq as new MSXML2.ServerXMLHTTP60

   httpReq.open "GET", "https://google.com", false
   httpReq.send

   if httpReq.Status <> 200 then
      debug.print ("I am so sorry")
      debug.print (httpReq.status & ": " & httpReq.statusText)
      debug.print (httpReq.responseText)
      exit sub
   end if

   debug.print (httpReq.responseText)

end sub ' }
