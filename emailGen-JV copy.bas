Option Explicit
Sub UpdatePromt()
Dim Prompt As String, SentFrom As String, CompName As String, CompContact As String
Dim VarName1 As String, VarName2 As String, VarName3 As String, VarName4 As String
Dim VarValue1 As String, VarValue2 As String, VarValue3 As String, VarValue4 As String
Dim SendTo As String, Tone As String
SentFrom = Admin.Range("D5").Value 'Sent From Name
CompName = Admin.Range("D6").Value  'Sent From Company
CompContact = Admin.Range("D7").Value  'Company Contact Info
With EmailGen
    Tone = .Range("O5").Value 'Email Tone
    SendTo = .Range("R5").Value 'Send to Name
    Prompt = .Range("H10").Value 'Prompt With Variables
    VarName1 = "[" & .Range("I7").Value & "]" 'Variable Name 1
    VarName2 = "[" & .Range("K7").Value & "]" 'Variable Name 2
    VarName3 = "[" & .Range("M7").Value & "]" 'Variable Name 3
    VarName4 = "[" & .Range("O7").Value & "]" 'Variable Name 4
    VarValue1 = .Range("R7").Value 'Variable Value 1
    VarValue2 = .Range("T7").Value 'Variable Value 2
    VarValue3 = .Range("V7").Value 'Variable Value 3
    VarValue4 = .Range("X7").Value 'Variable Value 4
    
    'Update Formats if not General
    If .Range("R7").NumberFormat <> "General" Then VarValue1 = Format(VarValue1, .Range("R7").NumberFormat)
    If .Range("T7").NumberFormat <> "General" Then VarValue2 = Format(VarValue2, .Range("T7").NumberFormat)
    If .Range("V7").NumberFormat <> "General" Then VarValue3 = Format(VarValue3, .Range("V7").NumberFormat)
    If .Range("X7").NumberFormat <> "General" Then VarValue4 = Format(VarValue4, .Range("X7").NumberFormat)
    
    'Replace from Email Settings
    Prompt = Replace(Replace(Replace(Replace(Replace(Prompt, "[Tone]", Tone), "[Send To]", SendTo), "[Sent From]", SentFrom), "[Company Name]", CompName), "[Company Contact]", CompContact)
    'Replace variables with values
    .Range("H23").Value = Replace(Replace(Replace(Replace(Prompt, VarName1, VarValue1), VarName2, VarValue2), VarName3, VarValue3), VarName4, VarValue4)
End With
End Sub
Sub Generate_Response()
    Dim oXMLHTTP As Object
    Dim Response As String, AIBody As String, Prompt As String, SecretKey As String, ParsedAnswer As String, Subject As String
    Dim Url As String, SendTo As String, SiteLink As String, Model As String, Temp As String, EmailType As String
    UpdatePromt 'Run Macro to Update prompt
    'Determine API Request Variables
    With Admin
        EmailType = .Range("G4").Value  'Email Type
        SecretKey = .Range("G6").Value 'Secret Key
        Model = .Range("G7").Value 'Model
        Temp = .Range("G8").Value 'Temperature
    End With
        If SecretKey = "" Then
            MsgBox "Please make sure to enter a secret key on the Admin sheet"
            Exit Sub
        End If
    With EmailGen
        Prompt = Replace(Replace(Replace(.Range("H23").Value, Chr(34), "\" & Chr(34)), vbNewLine, "\n"), Chr(10), "\n")  'Prompt
        If Prompt = "" Then
            MsgBox "Please make sure to enter a prompt"
            .Range("H10").Select
            Exit Sub
        End If
            .Shapes("GenTextMsg").Visible = msoCTrue 'Display Generating Response Shape
            Application.ScreenUpdating = True
            DoEvents
            Set oXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP")
            oXMLHTTP.Open "POST", "https://api.openai.com/v1/completions", False 'Set Open Posts For Text
            AIBody = "{""model"": """ & Model & """, ""prompt"": """ & Prompt & """, ""temperature"": " & Temp & " , ""max_tokens"": 3500}"   'Request Body For Text Response
            Debug.Print AIBody
            oXMLHTTP.SetRequestHeader "Content-Type", "application/json" 'Set Header Type
            oXMLHTTP.SetRequestHeader "Authorization", "Bearer " & SecretKey 'Set Header Authorization
            oXMLHTTP.Send AIBody
        
         Response = oXMLHTTP.ResponseText 'Extract Response
         Debug.Print Response
         If InStr(Response, "Incorrect API key provided") > 0 Then
            MsgBox "Incorrect API key provided. You can find your API key at " & vbCrLf & "https://platform.openai.com/account/api-keys."
            Exit Sub
         End If
         If InStr(Response, "maximum context length") > 0 Then
            MsgBox "The length of the prompt and response is too long, please break it up into smaller parts or reduce the temperature"
            Exit Sub
         End If
         ParsedAnswer = Right(Response, Len(Response) - InStrRev(Response, "[{" & Chr(34) & "text" & Chr(34) & ":" & Chr(34)) - 9) 'Remove First part of coded response
         ParsedAnswer = Right(ParsedAnswer, Len(ParsedAnswer) - 4) 'Remove First two empty lines
         ParsedAnswer = Left(ParsedAnswer, InStrRev(ParsedAnswer, Chr(34) & "," & Chr(34) & "index") - 1) 'Remove last part of coded response
         ParsedAnswer = Replace(Replace(ParsedAnswer, "\n", Chr(10)), "\" & Chr(34), Chr(34)) 'Replace new lines with carraige return characters and quotation marks
         
         'Extract Subject
         If InStr(ParsedAnswer, "Subject Line:") > 0 Then 'If Response contains 'Subject Line
                Subject = Replace(Left(ParsedAnswer, InStr(ParsedAnswer, Chr(10)) - 1), "Subject Line: ", "")
                ParsedAnswer = Replace(ParsedAnswer, "Subject Line: " & Subject & Chr(10), "")
         Else 'If Response only contains 'Subject:'
                Subject = Replace(Left(ParsedAnswer, InStr(ParsedAnswer, Chr(10)) - 1), "Subject: ", "")
                ParsedAnswer = Replace(ParsedAnswer, "Subject: " & Subject & Chr(10), "")
         End If
         
         
         If Left(ParsedAnswer, 1) = Chr(10) Then ParsedAnswer = Right(ParsedAnswer, Len(ParsedAnswer) - 1)  'Remove any beginning carriage returns if they exist
         .Range("R10").Value = Subject
        .Range("Q12").Value = ParsedAnswer
        .Shapes("GenTextMsg").Visible = msoFalse 'Hide Information text shape
     End With
End Sub
Sub Create_Email()
Dim OutApp As Object, OutMail As Object
Dim DispEmail As String, EmailType As String, Subject As String, Message As String, EmailTo As String
Dim EmailRow As Long
With EmailGen
    EmailTo = .Range("V5").Value 'Email to
    Subject = .Range("R10").Value 'Subject
    Message = .Range("Q12").Value 'Message
End With
DispEmail = Admin.Range("G5").Value 'Email Option (Yes/no)
    Set OutApp = CreateObject("Outlook.Application")
        Set OutMail = OutApp.CreateItem(0)
        With OutMail
            .To = EmailTo
            .Subject = Subject
            .Body = Message
            If DispEmail = "Yes" Then .Display Else .Send
        End With
        With EmailLog
            EmailRow = .Range("A99999").End(xlUp).Row + 1
            .Range("A" & EmailRow).Value = Now 'Set Day and time
            .Range("B" & EmailRow).Value = EmailTo 'Sent To
            .Range("C" & EmailRow).Value = Subject 'Email Subject
            .Range("D" & EmailRow).Value = Message ' Email Message
            .Range(EmailRow & ":" & EmailRow).WrapText = False
        End With
End Sub


'Draft of Create Email With Stationary
'Sub Create_Email()
'    Dim OutApp As Object, OutMail As Object
'    Dim DispEmail As String, EmailType As String, Subject As String, Message As String, EmailTo As String
'    Dim EmailRow As Long
'
'    With EmailGen
'        EmailTo = .Range("V5").Value 'Email to
'        Subject = .Range("R10").Value 'Subject
'        Message = .Range("Q12").Value 'Message
'    End With
'    DispEmail = Admin.Range("G5").Value 'Email Option (Yes/no)
'
'    Set OutApp = CreateObject("Outlook.Application")
'    Set OutMail = OutApp.CreateItem(0)
'
'    ' Set the stationary here
'    OutMail.GetInspector.WordEditor.Application.ActiveDocument.GetStationeryByName ("MyStationary")
'
'    With OutMail
'        .To = EmailTo
'        .Subject = Subject
'        .Body = Message
'        If DispEmail = "Yes" Then .Display Else .Send
'    End With
'
'    With EmailLog
'        EmailRow = .Range("A99999").End(xlUp).Row + 1
'        .Range("A" & EmailRow).Value = Now 'Set Day and time
'        .Range("B" & EmailRow).Value = EmailTo 'Sent To
'        .Range("C" & EmailRow).Value = Subject 'Email Subject
'        .Range("D" & EmailRow).Value = Message ' Email Message
'        .Range(EmailRow & ":" & EmailRow).WrapText = False
'    End With
'End Sub


