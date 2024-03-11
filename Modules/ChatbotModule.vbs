Attribute VBA_ModuleType=VBAModule
Sub ChatbotModule
' Developed by Halil Emre Yildiz (Github @JahnStar)
'#API Request ********************************************************************************************************************
Function PostApiRequest(messages)
    Dim request, key, endpoint, model, temperature
    key = "sk-yZgbUF7Ra5pT9jVYblFR83Bl9kFJziFTkqerW11tSYv7Z2MX"
    endpoint = "https://api.openai.com/v1/chat/completions"
    'endpoint = "http://localhost:5001/v1/chat/completions"
    model = "gpt-3.5-turbo"
    temperature = 0.7
    Set request = CreateObject("Microsoft.XMLHTTP")
'
    request.Open "POST", endpoint, False
    request.setRequestHeader "Content-Type", "application/json"
    request.setRequestHeader "Authorization", "Bearer " & key
    Dim requestText
    requestText = Replace("{""model"": """ & model & """, ""messages"": [" & WrapMessages(messages, True) & "], ""temperature"": " & temperature & "}", vbCrLf, "\n")
    request.Send requestText
'
    If request.Status = 200 Then
        PostApiRequest = ParseJSON(request.responseText, "content")
    Else
        MsgBox ParseJSON(request.responseText, "message"), vbCritical
        InputBox "", "", "debug: " & requestText
    End If
End Function
Function ParseJSON(jsonString, key)
    Dim startPos, endPos, keyPos, valueStartPos, valueEndPos
    Dim keyValue, valueStr

    ' Replace escaped characters
    jsonString = Replace(jsonString, "\""", "'")
    jsonString = Replace(jsonString, "\\", "\")
    jsonString = Replace(jsonString, "\n\n", vbCrLf)
    jsonString = Replace(jsonString, "\n", vbCrLf)

    startPos = InStr(jsonString, """" & key & """") ' Start position of the switch
    keyPos = InStr(startPos, jsonString, ":") ' Position of the ":" character of the key

    If keyPos > 0 Then
        valueStartPos = InStr(keyPos, jsonString, """") + 1 ' Start position of the value
        valueEndPos = InStr(valueStartPos, jsonString, """") ' End position of value

        valueStr = Mid(jsonString, valueStartPos, valueEndPos - valueStartPos) ' Value string
        ParseJSON = valueStr ' Return value
    Else
        ParseJSON = "Sorry, my mistake." ' Return empty string if key not found
    End If
End Function
'********************************************************************************************************************
Sub AskAI()
    On Error GoTo ErrorHandler
        Dim requestText, responseText
        requestText = GetText("TextBox1")
        '
        Call SetText("TextBox1", "", True)
        Call DataModule.AddMessage("user", requestText)
        
        responseText = PostApiRequest(DataModule.GetMessages)
        Call DataModule.AddMessage("assistant", responseText)
        '
        Call SetText("_output", responseText, False) ' Call SetText("_output", WrapMessages(DataModule.GetMessages, False), False)
        Call Update_Text_to_Display
        ' Call Speak(responseText)
        ' Draw Image
        Dim drawPrompt, stylePrompt
            If InStr(responseText, "_art(") > 0 Then
            drawPrompt = Split(responseText, "_art(")(1)
            drawPrompt = Replace(drawPrompt, ": ", ":")
            stylePrompt = Split(drawPrompt, ":")(1)
            stylePrompt = Split(stylePrompt, ")")(0)
            drawPrompt = Split(drawPrompt, ":")(0)
            Call DiffusionModule.InsertImageFromAI(drawPrompt, stylePrompt)
            Call SetShapeVisibility("_picture", True)
            Call SetText("_output", "", False)
            Exit Sub
        Else
            Call SetShapeVisibility("_picture", False)
        End If
        '
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbCritical
    Call ClearMessages
End Sub
Sub ClearMessages()
    DataModule.InitializeMessages ("")
    Call SetText("_output", "I’M JAHNVIS", False)
    Call SetShapeVisibility("_picture", False)
    On Error GoTo ErrorHandler
    Call SetText("TextBox1", "", True)
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbCritical
End Sub
Function GetLastLine(str)
    Dim lines
    lines = Split(str, vbCrLf)
    GetLastLine = lines(UBound(lines))
End Function
Sub Speak(text)
    Dim sapi
    On Error GoTo ErrorHandler
        Set sapi = CreateObject("sapi.spvoice")
        Set sapi.Voice = sapi.GetVoices.Item(0)
        sapi.Speak text
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbCritical
End Sub
' ActiveX ***********************************************************************************************************
Function GetText(objectName)
    GetText = ActivePresentation.Slides(ActivePresentation.SlideShowWindow.View.Slide.SlideIndex).Shapes(objectName).OLEFormat.Object.text
End Function
Sub SetText(shapeName, text, isObject)
If isObject = True Then
    ActivePresentation.Slides(ActivePresentation.SlideShowWindow.View.Slide.SlideIndex).Shapes(shapeName).OLEFormat.Object.text = text
Else
    Dim shape As shape
    Set shape = ActivePresentation.Slides(ActivePresentation.SlideShowWindow.View.Slide.SlideIndex).Shapes(shapeName)
    shape.TextFrame.TextRange.text = text
    shape.TextFrame2.AutoSize = ppAutoSizeShapeToFitText
    shape.TextFrame.TextRange.text = text
End If
End Sub
Sub SetShapeVisibility(shapeName, isVisible)
    Dim shape As shape
    Set shape = ActivePresentation.Slides(ActivePresentation.SlideShowWindow.View.Slide.SlideIndex).Shapes(shapeName)

    If isVisible Then
        shape.Visible = msoTrue
    Else
        shape.Visible = msoFalse
    End If
End Sub
'********************************************************************************************************************
Sub example()
    Dim messages() As Variant
    DataModule.GetMessages
    messages = DataModule.GetMessages
    '
    Call DataModule.AddMessage("user", "*Let's role play. You are a game developer and I am your best friend.* How are you?")
    Call DataModule.AddMessage("assistant", PostApiRequest(messages))
    MsgBox DataModule.WrapMessages(messages, False)
End Sub
Sub Update_Text_to_Display()
    Dim oSl    As Slide
    Dim slID   As Long
    Dim text2D As String
    Dim newText2D As String
    Dim otxR   As TextRange
    Dim pHyperlink As Hyperlink
    On Error Resume Next
    oSl = ActivePresentation.Slides(ActivePresentation.SlideShowWindow.View.Slide.SlideIndex)
    For Each pHyperlink In oSl.Hyperlinks
        If Val(pHyperlink.SubAddress) > 255 Then        ' SlideID
            text2D = pHyperlink.TextToDisplay
            slID = CLng(left(pHyperlink.SubAddress, InStr(pHyperlink.SubAddress, ",") - 1))
            newText2D = ActivePresentation.Slides.FindBySlideID(slID).SlideIndex
            If pHyperlink.Type = msoHyperlinkRange Then
                Set otxR = pHyperlink.Parent.Parent
            ElseIf pHyperlink.Type = msoHyperlinkShape Then
                Set otxR = pHyperlink.Parent.Parent.TextRange
            End If
            Set otxR = otxR.Replace(text2D, newText2D)
            Set otxR = otxR.Find(newText2D)
            With otxR.ActionSettings(ppMouseClick)
                .Action = ppActionHyperlink
                .Hyperlink.SubAddress = slID & "," & newText2D & "," & "Slide" & newText2D
            End With
        End If
    Next        'hyperlink
End Sub

End Sub