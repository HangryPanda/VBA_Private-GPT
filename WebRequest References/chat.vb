Option Explicit

Sub Chat()
    Dim lastQuestionRef As String
    Dim isLoading As Boolean
    Dim showLoadingMessage As Boolean
    Dim activeCitation As Variant
    Dim isCitationPanelOpen As Boolean
    Dim answers() As Variant
    Dim showAuthMessage As Boolean
    Dim initialData As Object
    
    Dim xmlhttp As Object
    Dim abortFuncs() As Object
    
    ' Create an instance of the XMLHTTP object
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP60")
    ' Create an array to store abort controllers
    ReDim abortFuncs(0)
    
    Dim i As Long
    Dim url As String
    Dim response As String
    
    Dim userMessage As Variant
    Dim request As Variant
    Dim result As Variant
    
    ' Set initial values
    lastQuestionRef = ""
    isLoading = False
    showLoadingMessage = False
    activeCitation = Empty
    isCitationPanelOpen = False
    ReDim answers(0)
    showAuthMessage = True
    
    ' Get the initial data
    Set initialData = Window("initialData")
    
    ' Other variables and declarations...
    
    ' Rest of the code goes here...
    
    ' Clean up
    Set xmlhttp = Nothing
    Erase abortFuncs
End Sub

Sub MakeApiRequest(question As String)
    ' Rest of the code goes here...
End Sub

Sub ClearChat()
    ' Rest of the code goes here...
End Sub

Sub StopGenerating()
    ' Rest of the code goes here...
End Sub

Sub OnShowCitation(citation As Variant)
    ' Rest of the code goes here...
End Sub

Sub ParseCitationFromMessage(message As Variant)
    ' Rest of the code goes here...
End Sub

Sub AddDefaultSrc(e As Variant)
    ' Rest of the code goes here...
End Sub