option explicit

sub main() ' {

    dim httpReq as new winHTTP.winHTTPRequest

    httpReq.setTimeouts 20000, 20000, 20000, 20000

    httpReq.Open "GET", "http://renenyffenegger.ch/"

    debug.print (httpReq.Option(WinHttpRequestOption_UserAgentString))
  ' Mozilla/4.0 (compatible; Win32; WinHttp.WinHttpRequest.5)

  ' httpReq.option(winHttpRequestOption_UserAgentString)              = "...."
    httpReq.option(winHttpRequestOption_EnableRedirects)              = true
    httpReq.option(winHttpRequestOption_EnableHttpsToHttpRedirects)   = true
  ' httpReq.option(winHttpRequestOption_SslErrorIgnoreFlags)          = 13056 ' Ignore all errors
  ' httpReq.option(winHttpRequestOption_EnablePassportAuthentication) = true

    httpReq.send

    if err.number <> 0 then ' {
       debug.print ("Could not perform http request")
       debug.print (err.Number & ": " & Err.Source)
       debug.print (err.Description)
       exit sub
    end if ' }

    if httpReq.Status <> 200 then
       debug.print (httpReq.Status & ": " & httpReq.StatusText)
    end if

    debug.print (httpReq.responseText)

end sub ' }
