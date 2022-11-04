Function GETENV(variableName)
    
    Set objWshShell     = WScript.CreateObject("WScript.Shell")
    Set WshSysEnv       = objWshShell.Environment("PROCESS")
    GETENV              = WshSysEnv(variableName)
    Set objWshShell     = Nothing

End Function

Function Dec2HexColor(decColor)
    If decColor > 16777215 Then decColor = 16777215
    If decColor < 0 Then decColor = 0
    Dec2HexColor = "#" & Right("00" & Hex(decColor Mod 256), 2) & _
                         Right("00" & Hex((decColor \ 256) Mod 256), 2) & _
                         Right("00" & Hex(decColor \ 65536), 2)
End Function

Function ActivateOutlook()

    Set objWshShell     = WScript.CreateObject("WScript.Shell")
    Set objDictionary   = CreateObject("Scripting.Dictionary")
    Set objServices     = GetObject("winmgmts:\\.\root\CIMV2")
    Set objProcess      = objServices.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'outlook.exe'")

    For Each Process In objProcess
        objWshShell.AppActivate Process.ProcessId
    Next

    Set objProcess      = Nothing
    Set objDictionary   = Nothing
    Set objWshShell     = Nothing
    
End Function

Dim objApp, objMail

On Error Resume Next
Set objApp = GetObject("","Outlook.Application")
If objApp Is Nothing Then
    Set objApp = CreateObject("Outlook.Application")
End If
Err.Clear

If Not objApp Is Nothing Then
    On Error Resume Next
    Set objMail = objApp.CreateItem(0)
    Err.Clear
End If

If Not objMail Is Nothing Then
  
    objMail.To      = GETENV("MAIL_TO")
    objMail.Cc      = GETENV("MAIL_CC")
    objMail.Bcc     = GETENV("MAIL_BCC")
    objMail.Subject = GETENV("MAIL_SUBJECT")
    objMail.BodyFormat  = 2
    
    text = WScript.StdIn.ReadAll
    
    If GETENV("MAIL_FORMAT") Is "TEXT" Then
    
        objMail.Body        = " "
        
        Set objInspector = objMail.GetInspector
        Set objWordEditor = objInspector.WordEditor
        
        If Not (objWordEditor Is Nothing) Then
        
            text_length = objWordEditor.Characters.Count
            
            Set objRange = objWordEditor.Range(0, text_length)
            Set objFont = objRange.Font
            
            font_name   = objFont.Name
            font_size   = objFont.Size
            font_italic = objFont.Italic
            font_bold   = objFont.Bold
            font_color  = Dec2HexColor(objFont.Color)
                                    
            style = "font-family:'" & font_name & "';color:" & font_color & ";font-size:" & font_size & ";"
            
            If font_italic Then
                style = style & "font-style:italic;"
            End If
            If font_bold Then
                style = style & "font-weight:bold;"
            End If

            objMail.HTMLBody = "<span style=" & Chr(34) & style & Chr(34) & ">" & text & "</span>"
            text_end = objWordEditor.Range.End-1
            Set objRange = objWordEditor.Range(text_end, text_end)
            objRange.Select
            
            Set objRange        = Nothing
            Set objFont         = Nothing
            Set objWordEditor   = Nothing
            
        End If
        
        Set objInspector = Nothing
        
    Else
    
        objMail.HTMLBody    = text
        
        Set objInspector = objMail.GetInspector
        Set objWordEditor = objInspector.WordEditor
        
        If Not (objWordEditor Is Nothing) Then
            text_end = objWordEditor.Range.End-1
            Set objRange = objWordEditor.Range(text_end, text_end)
            objRange.Select
            Set objRange        = Nothing
            Set objWordEditor   = Nothing
        End If
        
        Set objInspector = Nothing
        
    End If
    
    attachmentCount = 1
    
    Do While 1
        attachment = GETENV("MAIL_ATTACHMENT_" & attachmentCount)
        If attachment <> "" Then
            objMail.Attachments.Add attachment
            attachmentCount = attachmentCount + 1
        Else
            Exit Do
        End If
    Loop

    objMail.Display
    
    Set objMail = Nothing
        
End If

ActivateOutlook()
