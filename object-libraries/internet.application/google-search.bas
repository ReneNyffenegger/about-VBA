'  Needed references:
'      Microsoft Internet Controls
'         c:\Windows\...\ieframe.dll
'
'         call application.workbooks(1).VBProject.references.addFromGuid("{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}")
'        (Name is: SHDocVw)
'
'     Microsoft HTML Object Library
'         c:\Windows\...\mshtml.tlb
'         call application.workbooks(1).VBProject.references.addFromGuid("{3050F1C5-98B5-11CF-BB82-00AA00BDCE0B}")
'        (Name is: MSHTML)
'

option explicit

sub main()
    dim ie           as new internetExplorer ' set ie = createObject("InternetExplorer.Application")
    dim doc          as     htmlDocument
    dim inputSearch  as     htmlInputElement
    dim buttonSubmit as     htmlButtonElement

    ie.visible = true
    ie.navigate "https://www.google.com"

  ' Wait for page to be loaded
    do while ie.busy
       application.wait dateAdd("s", 1, now)
    loop

    set doc          = ie.document

  ' Get Elements on page:
    set inputSearch  = doc.querySelector("input[name='q']"   )
    set buttonSubmit = doc.querySelector("input[name='btnK']")

  ' Fill input box for query
    inputSearch.value = "internet.application"

  ' Click button
    buttonSubmit.click

end sub
