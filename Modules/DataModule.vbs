Attribute VBA_ModuleType=VBAModule
Sub DataModule
' Developed by Halil Emre Yildiz (Github @JahnStar)
' DataModule
Option Explicit
Public SharedMessages() As Variant
Private Initialized As Boolean
Function GetMessages() As Variant
    If Not Initialized Then
        InitializeMessages ("")
    End If
    If Initialized Then
        GetMessages = SharedMessages
        Exit Function
    End If
    MsgBox "Error"
End Function
Sub InitializeMessages(prompt)
    SharedMessages = Array(Array("system", "[You are an AI named JAHNVIS, in a presentation that answers in a word, developed by Halil Emre Yildiz (AKA Jahn Star).] You can draw whit this function generate_art(englishPrompt:style) englishPrompt parameter must be in English. Must be use ':' to seperate paramters. style paramter must be one of these (photographic, digital-art, anime, fantasy-art, cinematic) Example conversation 'User: Hello there!, AI: Hi. User: How are you? AI: Good. User: Who are you? AI: I'm JAHNVIS User: Can AI make art? AI: Wht not? User: So, can you draw? AI: Yes. User: So, suprise me! AI: Sure. [generate_art(a cyberpunk girl, bodysuit, neon color, multicolor hair, science fiction, close up:anime)] User: Wow, that's Amazing! AI: Thanks. User: Can you draw a surreal floating island? AI: Sure. [generate_art(urrealist painting of a floating island with giant clock gears, populated with mythical creatures:fantasy-art)]'" & prompt))
    Initialized = True
End Sub
Sub AddMessage(role, message)
    If Not Initialized Then
        InitializeMessages ("")
    End If
    ReDim Preserve SharedMessages(UBound(SharedMessages) + 1)
    SharedMessages(UBound(SharedMessages)) = Array(role, message)
End Sub
Function WrapMessages(messages, toJson)
    Dim json, i
    For i = 0 To UBound(messages)
        Dim role, content
        role = messages(i)(0)
        content = messages(i)(1)

        If (toJson = True) Then
            json = json & "{""role"": """ & role & """, ""content"": """ & content & """}"
            If i < UBound(messages) Then
                json = json & ", "
            End If
        Else
            If role <> "system" Then
                If role <> "user" Then
                    role = "- (AI) "
                Else
                    role = "- You: "
                End If
                json = json & role & content & vbCrLf
            End If
        End If
    Next
    WrapMessages = json
End Function

End Sub
