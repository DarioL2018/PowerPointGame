Attribute VB_Name = "Módulo1"
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" _
      Alias "URLDownloadToFileA" ( _
        ByVal pCaller As LongPtr, _
        ByVal szURL As String, _
        ByVal szFileName As String, _
        ByVal dwReserved As LongPtr, _
        ByVal lpfnCB As LongPtr _
      ) As Long
      
Public Sub descargarImagenes()
recoverImages
InsertarImagenes

'Dim myurl As String
'Dim xmlhttp As Object 'New MSXML2.XMLHTTP60
'Dim oHtml       As HTMLDocument
'Dim oElement    As Object
'Dim index As Integer
'Dim dlpath As String


'index = 0
'Set xmlhttp = CreateObject("MSXML2.serverXMLHTTP")
'Set xmlhttp = CreateObject("WinHTTP.WinHTTPRequest.5.1")
'myurl = "https://www.google.com/search?as_st=y&tbm=isch&as_q=puppy&as_epq=&as_oq=&as_eq=&cr=&as_sitesearch=&safe=images&tbs=isz:l,ift:jpg,sur:fmc"
'xmlhttp.Open "GET", myurl, False
'xmlhttp.setRequestHeader "Content-Type", "text/xml"
'xmlhttp.setRequestHeader "DNT", "1"
'xmlhttp.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; rv:25.0) Gecko/20100101 Firefox/25.0"
'xmlhttp.send
'MsgBox (xmlhttp.responseText)
'Dim stringa() As String

'Set oHtml = New HTMLDocument
'oHtml.body.innerHTML = xmlhttp.responseText

'For Each oElement In oHtml.getElementsByClassName("rg_ic rg_i")
'URLDownloadToFile 0, oElement.src, dlpath & "test.jpg", 0, 0
'If (InStr(1, oElement.src, "jpeg") > 0) Then
'dlpath = "C:\tmp\tempo.jpeg"
'stringa = Split(oElement.src, ";base64,")
'Open dlpath For Binary As #1
'           Put #1, 1, DecodeBase64(stringa(1))
'        Close #1
'Exit For
'End If
'Next

End Sub
Private Function DecodeBase64(ByVal strData As String) As Byte()

    Dim objXML As Object 'MSXML2.DOMDocument
    Dim objNode As Object 'MSXML2.IXMLDOMElement

    'get dom document
    Set objXML = CreateObject("MSXML2.DOMDocument")

    'create node with type of base 64 and decode
    Set objNode = objXML.createElement("b64")
    objNode.dataType = "bin.base64"
    objNode.Text = strData
    DecodeBase64 = objNode.nodeTypedValue

    'clean up
    Set objNode = Nothing
    Set objXML = Nothing

End Function
Function recoverImages()
    'Dim miArray(5) As String
    Dim url2 As String
    Dim div, div2 As HTMLDivElement
    Const ERROR_SUCCESS As Long = 0
    Const BINDF_GETNEWESTVERSION As Long = &H10
    Const INTERNET_FLAG_RELOAD As Long = &H80000000
    Const folderName As String = "c:\temp\"
    Dim htmlPrincipal As Object, htmlPrincipal2 As Object
    Dim xmlhttp As Object
    Dim ws As MSXML2.xmlhttp
    Dim cuerpo As HTMLBody
    Dim docum As HTMLDocument, docum2 As HTMLDocument
    Dim url4 As String
    
    'ws.responseBody
    Set xmlhttp = CreateObject("MSXML2.xmlHttp")
    
    'my_url = "https://www.google.com/search?as_st=y&tbm=isch&as_q=puppy&as_epq=&as_oq=&as_eq=&cr=&as_sitesearch=&safe=images&tbs=isz:l,ift:jpg,sur:fmc"
    'my_url = "https://www.bing.com/images/search?q=puppy&qs=n&form=QBIR&qft=%20filterui%3Aimagesize-large&sp=-1&pq=&sc=0-0&sk=&cvid=D8F17019D2AA4DBAB81576CC9043B80C"
    
    With xmlhttp
        .Open "GET", "https://www.bing.com/images/search?q=chocolate&qft=+filterui:imagesize-large+filterui:license-L2_L3", False
        .send
        resp = .responseText
    End With
    

    Set htmlPrincipal = CreateObject("HTMLFile")
    htmlPrincipal.write resp
    Set docum = htmlPrincipal.body.document
    'url4 = htmlPrincipal.body.document.getElementsByClassName("iusc")(0).href
    
    url4 = docum.getElementsByClassName("iusc")(0).href
    urlImagen = getUrl(url4)
    dlpath = "C:\tmp\imagen.jpg"
    urlImagen2 = URLDecode(urlImagen)
    Debug.Print urlImagen2
    URLDownloadToFile 0&, urlImagen2, dlpath, BINDF_GETNEWESTVERSION, 0&
    'With xmlhttp
    '    .Open "GET", url4, False
    '    .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    '    .setRequestHeader "DNT", "1"
    '    .setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.1; rv:25.0) Gecko/20100101 Firefox/25.0"
    '    .send
    '    resp2 = .responseText
    'End With
    
    'Set htmlPrincipal2 = CreateObject("HTMLFile")
    'htmlPrincipal2.write resp2
    'Set docum2 = htmlPrincipal2.body.document
    'Set div2 = docum2.getElementsByClassName("imgContainer nofocus")(0)
    ' dlpath = "C:\tmp\tempo.html"
    'URLDownloadToFile 0&, url4, dlpath, BINDF_GETNEWESTVERSION, 0&
    
    'urlf2 = div2.getElementsByTagName("img")(0).src
    'End With
    
    
    'my_url = "https://www.bing.com/images/search?q=cat&qft=+filterui:imagesize-large+filterui:license-L2_L3"
    Set ie = CreateObject("InternetExplorer.Application")
    Debug.Print url4
    With ie
        .Visible = True
       '.navigate my_url
        .navigate url4
        '.Top = 50
        '.Left = 530
        '.Height = 400
        '.Width = 400
    End With

'    Do Until Not ie.Busy And ie.readyState = 4
'        DoEvents
'    Loop
    
    dlpath = "C:\tmp\tempo.jpeg"
    'stringa = Split(ie.document.getElementsByClassName("rg_ic rg_i")(1).src, ";base64,")
'    url2 = ie.document.getElementsByClassName("iusc")(0).href
    
    
    With ie
        .Visible = False
        .navigate url4
        '.Top = 50
        '.Left = 530
        '.Height = 400
        '.Width = 400
    End With

    Do Until Not ie.Busy And ie.readyState = 4
        DoEvents
    Loop
    dlpath = "C:\tmp\tempo.jpeg"
    
    
    '"https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcRHFJtmkS0i6-Uu48UZYQYPwHv4p9I64mRmtrzRUqj3-5y6kHmiag&s"
    Set div = ie.document.getElementsByClassName("imgContainer nofocus")(0)
    urlf = div.getElementsByTagName("img")(0).src
    
    URLDownloadToFile 0&, urlf, dlpath, BINDF_GETNEWESTVERSION, 0&
    
    'Open dlpath For Binary As #1
    'Put #1, 1, DecodeBase64(stringa(1))
    'Close #1
    
End Function

Function OnReadyStateChange()
    
        MsgBox "Done"
    
End Function

Sub InsertarImagenes() '(titulo As String, imagenes As Object)
Dim archivoPPT As PowerPoint.Application
Dim diapositiva As PowerPoint.Slide
Dim tablaTotal() As Shape
Dim largo, ancho, dimension As Integer
largo = 10
ancho = 10
dimension = 50
ReDim tablaTotal(0 To largo, 0 To ancho)

'Dim coll As Object
'Set coll = CreateObject("System.Collections.ArrayList")
    
'Instancia del objeto PowerPoint.Application
Set archivoPPT = New PowerPoint.Application
 
'Creamos una presentación PowerPoint
archivoPPT.Presentations.Add
 
'Recorrer tods los gráficos en nuestro libro de Excel

Dim item As Variant

'For Each item In imagenes
'     Debug.Print item
'Next item
         
        'Agregar nueva diapositiva
        archivoPPT.ActivePresentation.Slides.Add _
            archivoPPT.ActivePresentation.Slides.Count + 1, ppLayoutBlank
        archivoPPT.ActiveWindow.View.GotoSlide _
            archivoPPT.ActivePresentation.Slides.Count
        Set diapositiva = archivoPPT.ActivePresentation.Slides( _
            archivoPPT.ActivePresentation.Slides.Count)
                 
        'Copiar gráfico en la dispositiva
        
        ActiveWindow.Selection.SlideRange.Shapes.AddPicture( _
        FileName:="C:\tmp\tempo.jpeg", _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, Left:=60, Top:=35, _
        Width:=98, Height:=48).Select
         
        'Centramos la imagen insertada
        archivoPPT.ActiveWindow.Selection.ShapeRange.Align msoAlignCenters, msoTrue
        archivoPPT.ActiveWindow.Selection.ShapeRange.Align msoAlignMiddles, msoTrue
         
        Dim sh As Shape
        For i = 0 To largo
        For j = 0 To ancho
            Set sh = ActiveWindow.Selection.SlideRange.Shapes.AddShape(Type:=msoShapeRectangle, _
    Left:=j * dimension, Top:=i * 50, Width:=dimension, Height:=dimension)
            'Set tablaTotal(i, j) = sh
        Next j
        Next i
              
        sh.Fill.ForeColor.RGB = RGB(220, 105, 0)
        sh.Delete
        
    
'Eliminamos las instancias creadas
Set diapositiva = Nothing
Set archivoPPT = Nothing

End Sub

Private Function URLDecode(ByVal txt As String) As String
Dim txt_len As Integer
Dim i As Integer
Dim ch As String
Dim digits As String
Dim result As String

    'SetSafeChars

    result = ""
    txt_len = Len(txt)
    i = 1
    Do While i <= txt_len
        ' Examine the next character.
        ch = Mid$(txt, i, 1)
        If ch = "+" Then
            ' Convert to space character.
            result = result & " "
        ElseIf ch <> "%" Then
            ' Normal character.
            result = result & ch
        ElseIf i > txt_len - 2 Then
            ' No room for two following digits.
            result = result & ch
        Else
            ' Get the next two hex digits.
            digits = Mid$(txt, i + 1, 2)
            result = result & Chr$(CInt("&H" & digits))
            i = i + 2
        End If
        i = i + 1
    Loop

    URLDecode = result
End Function

Function getUrl(substring As String) As String
    Dim startSymbol As Integer 'mediaurl=
    Dim endSymbol As Integer   '&
    Dim result As String
    Dim LArray() As String
 
    result = ""
 
    startSymbol = InStr(substring, "mediaurl=") + 9
    endSymbol = InStr(startSymbol, substring, "&")
 
    If startSymbol > 0 And endSymbol > startSymbol Then
        result = Mid(substring, startSymbol, endSymbol - startSymbol)
    Else
        result = substring
    End If
    getUrl = result
End Function


