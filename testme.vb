'Option Explicit
'Microsoft Internet Controls

Public Sub testme()
    Dim IE As InternetExplorer

    Set IE = New InternetExplorer
    With IE
        .navigate "http://www.vitoshacademy.com/archives/"
        .Visible = False
        While .Busy Or .readyState <> READYSTATE_COMPLETE
            DoEvents
        Wend
        Set objHTML = .document
        DoEvents
    End With
    Set elementONE = objHTML.getElementsByTagName("li")
    For i = 1 To elementONE.Length
        elementTwo = elementONE.Item(i).innerText
        Debug.Print elementTwo
    Next i

    DoEvents
    IE.Quit
    DoEvents
    Set IE = Nothing

End Sub
