Attribute VB_Name = "Módulo1"
Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" _
      Alias "URLDownloadToFileA" ( _
        ByVal pCaller As LongPtr, _
        ByVal szURL As String, _
        ByVal szFileName As String, _
        ByVal dwReserved As LongPtr, _
        ByVal lpfnCB As LongPtr _
      ) As Long
Dim path As String
Dim widthSquare, heightSquare, leftSquare, topSquare
Public Const dimension As Integer = 90

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
Sub downloadAndRecoverImages(keyword As String)

    Const BINDF_GETNEWESTVERSION As Long = &H10
    Dim htmlPrincipal As Object
    Dim xmlhttp As Object
    Dim docum As HTMLDocument
    Dim url4 As String
    Dim intCode As Long
    
    Set xmlhttp = CreateObject("MSXML2.xmlHttp")
    'my_url = "https://www.google.com/search?as_st=y&tbm=isch&as_q=puppy&as_epq=&as_oq=&as_eq=&cr=&as_sitesearch=&safe=images&tbs=isz:l,ift:jpg,sur:fmc"
    With xmlhttp
        .Open "GET", "https://www.bing.com/images/search?q=" + keyword + "&qft=+filterui:imagesize-large+filterui:license-L2_L3+filterui:photo-photo", False
        .send
        resp = .responseText
    End With
    
    Set htmlPrincipal = CreateObject("HTMLFile")
    htmlPrincipal.write resp
    Set docum = htmlPrincipal.body.document
    For i = 0 To 4
        dlpath = path & "\" & keyword & "_" & (i + 1) & ".jpg"
        url4 = docum.getElementsByClassName("iusc")(i).href
        urlImagen = URLDecode(getParameter(url4, "mediaurl"))
        Debug.Print url4
        'exph=996&expw=1024
        intCode = URLDownloadToFile(0&, urlImagen, dlpath, BINDF_GETNEWESTVERSION, 0&)
        If (intCode = 0) Then
            InsertarImagenes keyword, (dlpath)
            Kill (dlpath)
        End If
    Next
    
End Sub

Sub mainProgram()
    Dim arrKeyWords As Object
    Set arrKeyWords = CreateObject("System.Collections.ArrayList")
    path = ActivePresentation.path
    
    Set arrKeyWords = readFile()
    
    For Each word In arrKeyWords
        downloadAndRecoverImages (word)
    Next
    ActivePresentation.SaveAs path & "\" & "Vocabulary.pptx"
    cubrirImagenes
    ActivePresentation.SaveAs path & "\" & "Hidden Pictures", ppSaveAsOpenXMLPresentationMacroEnabled
    MsgBox "Successful execution!"
    ActivePresentation.Application.Quit
    
End Sub

Function readFile() As Object
    Dim myFile As String
    Dim textline As String
    Dim arrKeyWords As Object
    
    texline = ""
    myFile = path & "\" & "words.txt"
    Set arrKeyWords = CreateObject("System.Collections.ArrayList")
    'Open Plain text File
    Open myFile For Input As #1
    
    'Read File
    Do Until EOF(1)
        Line Input #1, textline
        'search user
        If Len(textline) > 0 Then
            arrKeyWords.Add textline
            Debug.Print textline
        End If
    Loop
    Close #1
    Set readFile = arrKeyWords
End Function
Sub InsertarImagenes(titulo As String, pathImage As String)

    Dim archivoPPT As PowerPoint.Application
    Dim diapositiva As PowerPoint.Slide
    Dim tablaTotal() As Shape
    
    Dim oSlides As Slides, oSlide As Slide
    Set oSlides = ActivePresentation.Slides
    
    Set oSlide = oSlides.AddSlide(ActivePresentation.Slides.Count + 1, _
    GetLayout("SmileBlank"))
    oSlide.Select
    Dim item As Variant
        
    Dim imageA As Shape
    Set imageA = oSlide.Shapes.AddPicture( _
    FileName:=pathImage, _
    LinkToFile:=msoFalse, _
    SaveWithDocument:=msoTrue, Left:=0, Top:=0, _
    Width:=-1, Height:=-1)
    widthSquare = imageA.Width
    heightSquare = imageA.Height
    leftSquare = imageA.Left
    topSquare = imageA.Top
    oSlide.Shapes(2).TextFrame.TextRange.Text = UCase(titulo)
    oSlide.NotesPage.Shapes(2).TextFrame.TextRange.Text = UCase(titulo)

End Sub

Sub cubrirImagenes()
    Dim archivoPPT As PowerPoint.Application
    Dim diapositiva As PowerPoint.Slide
    Dim tablaTotal() As Shape
    Dim largo, ancho As Integer
     
    ancho = Round(widthSquare / dimension)
    largo = Round(heightSquare / dimension)
    Debug.Print "widthSquare: " & widthSquare & " Calculation: " & ancho
    Debug.Print "heightSquare: " & heightSquare & " Calculation: " & largo
    
    Dim oSlides As Slides, oSlide As Slide
    Set oSlides = ActivePresentation.Slides
    
    For Index = 1 To oSlides.Count
        oSlides(Index).CustomLayout = GetLayout("SmileCover")
         
        Dim sh As Shape
        For i = 0 To largo - 1
        For j = 0 To ancho - 1
            Set sh = oSlides(Index).Shapes.AddShape(Type:=msoShapeRectangle, _
    Left:=(leftSquare - 6 + (j * dimension)), Top:=(topSquare - 50 + (i * dimension)), Width:=dimension, Height:=dimension)
            sh.Fill.ForeColor.RGB = randColour
        Next j
        Next i
              
'        sh.Fill.ForeColor.RGB = RGB(220, 105, 0)
'        sh.Delete
        'oSlides(Index).Shapes(2).ZOrder msoBringToFront
    Next Index
End Sub

Public Function GetLayout( _
    LayoutName As String, _
    Optional ParentPresentation As Presentation = Nothing) As CustomLayout

    If ParentPresentation Is Nothing Then
        Set ParentPresentation = ActivePresentation
    End If

    Dim oLayout As CustomLayout
    For Each oLayout In ParentPresentation.SlideMaster.CustomLayouts
        If oLayout.Name = LayoutName Then
            Set GetLayout = oLayout
            Exit For
        End If
    Next
End Function

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

Function getParameter(substring As String, parameter As String) As String
    Dim startSymbol As Integer '
    Dim endSymbol As Integer   '&
    Dim result As String
    Dim LArray() As String
 
    result = ""
 
    startSymbol = InStr(substring, parameter & "=") + Len(parameter) + 1
    endSymbol = InStr(startSymbol, substring, "&")
 
    If startSymbol > 0 And endSymbol > startSymbol Then
        result = Mid(substring, startSymbol, endSymbol - startSymbol)
    Else
        result = substring
    End If
    getParameter = result
End Function

Sub removeSquare()
    'cubrirImagenes
    Dim Shapes As Object
    Dim indexS As Integer
    
    Set Shapes = ActivePresentation.SlideShowWindow.View.Slide.Shapes
    
    indexS = 0
    'Debug.Print "Cantidad de Figuras" & Shapes.Count
    For i = 1 To Shapes.Count
         If (Shapes(i).Type = msoShapeRectangle) Then
            indexS = i
            Exit For
        End If
    Next i
    
    If indexS > 0 Then
        'Shapes(randNumber(indexS, Shapes.Count)).Visible = msoFalse
        Shapes(randNumber(indexS, Shapes.Count)).Delete
    Else
        ActivePresentation.SlideShowWindow.View.Slide.CustomLayout = GetLayout("SmileBlank")
    End If
End Sub

Sub removeAllSquares()
    'cubrirImagenes
    Dim ShapesAll As Shapes
   
    Set ShapesAll = ActivePresentation.SlideShowWindow.View.Slide.Shapes
    'Set ShapesAll = ActivePresentation.Slides(8).Shapes
    'Debug.Print "Cantidad de Figuras" & ShapesAll.Count
    For i = ShapesAll.Count To 1 Step -1
         If (ShapesAll(i).Type = msoShapeRectangle) Then
            ShapesAll(i).Delete
        End If
    Next i
    ActivePresentation.SlideShowWindow.View.Slide.CustomLayout = GetLayout("SmileBlank")
End Sub


Function randNumber(upperbound As Integer, lowerbound As Integer) As Integer
    randNumber = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
End Function

Function randColour()
    ConstantArray = Array(vbBlack, vbBlue, vbCyan, vbRed, vbGreen, vbYellow _
    , vbMagenta, vbWhite)
    
    randColour = ConstantArray(randNumber(1, 8))
End Function


