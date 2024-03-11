Attribute VBA_ModuleType=VBAModule
Sub DiffusionModule
' Developed by Halil Emre Yildiz (Github @JahnStar)
Sub InsertImageFromAI(prompt, style)
    Dim request, key, endpoint, size
    key = "sk-8tdom7s5qxym8Srsqedu1we7TdL7X2I1DAR1e4u4xynrf52M"
    endpoint = "https://api.stability.ai/v1/generation/stable-diffusion-v1-6/text-to-image"
    size = 512
    ' style = "enhance" ' enhance, anime, photographic, digital-art, comic-book, fantasy-art, 3d-model
    Set request = CreateObject("Microsoft.XMLHTTP")
    ' MsgBox prompt ' debug
    request.Open "POST", endpoint, False
    request.setRequestHeader "Content-Type", "application/json"
    request.setRequestHeader "Authorization", "Bearer " & key
    
    If style <> "photographic" And style <> "digital-art" And style <> "anime" And style <> "fantasy-art" And style <> "cinematic" Then
    style = "photographic"
    End If
    jsonString = "{ ""steps"": 20, ""width"": " & size & ", ""height"": " & size & ", ""seed"": 0, ""cfg_scale"": 6, ""samples"": 1, ""style_preset"": """ & style & """, ""text_prompts"": [ { ""text"": """ & prompt & """, ""weight"": 1 }, { ""text"": ""blurry, bad, worst quality, NSFW"", ""weight"": -1 } ] }"
    
    request.Send jsonString
    
    If request.Status = 200 Then
        Dim base64String, pngData
        base64String = Split(request.responseText, """")(5)
        
        ' base64 to byte array
        With CreateObject("Microsoft.XMLDOM").createElement("b64")
            .DataType = "bin.base64"
            .text = base64String
            pngData = .nodeTypedValue
        End With
        
        Dim filePath
        filePath = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%TEMP%") & "\output.png"
        'filePath = "output.png"
    
        ' byte array to PNG
        With CreateObject("ADODB.Stream")
            .Type = 1 ' binary
            .Open
            .Write pngData
            .SaveToFile filePath, 2 ' overwrite
            .Close
        End With

        Call InsertPicture("_picture", filePath)
    Else
        InputBox "", "", request.responseText, vbCritical
    End If
    Exit Sub
ErrorHandler:
    Call ChatbotModule.SetText("_output", "I draw when alone", False)
End Sub
Sub InsertPicture(shapeName, pictureURL)
    On Error GoTo ErrorHandler
    '--
    Dim Slide As Slide
    Dim picture As shape
    Dim url As String
    Dim left As Single
    Dim top As Single
    Dim width As Single
    Dim height As Single
    ' URL of the image
    url = pictureURL
    ' Get the first slide
    Set Slide = ActivePresentation.Slides(ActivePresentation.SlideShowWindow.View.Slide.SlideIndex)
    ' Find the picture with the name shapeName
    On Error Resume Next
    Set picture = Slide.Shapes(shapeName)
    If picture Is Nothing Then
        MsgBox "Picture not found"
        Exit Sub
    End If
    On Error GoTo 0
    ' Store the position and size of the picture
    left = picture.left
    top = picture.top
    width = picture.width
    height = picture.height
    ' Delete the picture
    picture.Delete
    ' Add the new picture
    Set picture = Slide.Shapes.AddPicture(url, msoFalse, msoTrue, left, top, width, height)
    ' Name the picture
    picture.Name = shapeName
    ' Add animation effect
    Dim effect As effect
    Set effect = Slide.TimeLine.MainSequence.AddEffect(picture, msoAnimEffectRandomBars, , msoAnimTriggerWithPrevious)
    effect.Timing.TriggerDelayTime = 2
    effect.Timing.Duration = 2
    '--
    Exit Sub
ErrorHandler:
    MsgBox Err.Description, vbCritical
End Sub

End Sub