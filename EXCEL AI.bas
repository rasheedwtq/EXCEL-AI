Attribute VB_Name = "Module1"
' Global variables for storing API configurations
Public g_APIKey As String
Public g_Model As String
Public g_APIEndpoint As String

' Worksheet name for storing configurations
Private Const CONFIG_SHEET As String = "_AIConfig_"

' Function to configure AI model and API key
Private Sub ConfigureAIModel(apiKey As String, Optional model As String = "gpt-3.5-turbo", Optional apiEndpoint As String = "https://api.openai.com/v1/chat/completions")
    g_APIKey = apiKey
    g_Model = model
    g_APIEndpoint = apiEndpoint
    
    MsgBox "AI model configured successfully!" & vbCrLf & _
           "Model: " & model & vbCrLf & _
           "API Endpoint: " & apiEndpoint, vbInformation
    
    ' Save configuration and test connection
    SaveAIConfig
    TestAIConnection
End Sub

' Function to test API connection
Private Sub TestAIConnection()
    If g_APIKey = "" Or g_Model = "" Or g_APIEndpoint = "" Then
        MsgBox "Please configure the API key, model name, and endpoint first!", vbExclamation
        Exit Sub
    End If
    
    Dim testResult As String
    testResult = AI_PROCESS(Selection, "Please reply with 'Connection successful'")
    
    If InStr(testResult, "Connection successful") > 0 Then
        MsgBox "Test result: Connection successful!", vbInformation
    Else
        MsgBox "Test result: Connection failed! " & testResult, vbExclamation
    End If
End Sub

' Main function - Process single cell or range
Function AI_PROCESS(cellData As Range, prompt As String) As String
    On Error GoTo ErrorHandler
    
    ' Check if API key is configured
    If g_APIKey = "" Then
        AI_PROCESS = "Error: Please configure the API key first"
        Exit Function
    End If
    
    ' Handle both single cell and range
    Dim dataText As String
    If cellData.Cells.Count = 1 Then
        ' Single cell
        dataText = CStr(cellData.Value)
    Else
        ' Range of cells
        dataText = ""
        Dim cell As Range
        For Each cell In cellData
            If IsNumeric(cell.Value) Then
                dataText = dataText & cell.Value & ", "
            Else
                dataText = dataText & CStr(cell.Value) & ", "
            End If
        Next cell
        ' Remove trailing comma and space
        If Len(dataText) > 2 Then
            dataText = Left(dataText, Len(dataText) - 2)
        End If
    End If
    
    ' Construct prompt for AI
    Dim fullPrompt As String
    fullPrompt = "Process the following instruction: " & prompt & " Data: " & dataText
    
    ' Call AI API
    AI_PROCESS = CallAI_API(fullPrompt)
    Exit Function
    
ErrorHandler:
    AI_PROCESS = "Error: " & Err.Description
End Function



' Core function to call AI API and simplify JSON parsing
Function CallAI_API(prompt As String) As String
    On Error GoTo ErrorHandler
    
    Dim httpRequest As Object
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    
    ' Prepare request data
    Dim postData As String
    postData = "{""model"": """ & g_Model & """, ""messages"": [{""role"": ""user"", ""content"": """ & EscapeJSON(prompt) & """}], ""temperature"": 0.7}"
    
    ' Send HTTP request
    httpRequest.Open "POST", g_APIEndpoint, False
    httpRequest.setRequestHeader "Content-Type", "application/json"
    httpRequest.setRequestHeader "Authorization", "Bearer " & g_APIKey
    httpRequest.send (postData)
    
    ' Parse response - Simplified JSON parsing
    If httpRequest.Status = 200 Then
        Dim response As String
        response = httpRequest.responseText
        
        ' Extract content field, avoiding complex JSON parsing
        CallAI_API = ExtractContentFromJSON(response)
    Else
        CallAI_API = "API error: " & httpRequest.Status & " - " & httpRequest.statusText & " - " & httpRequest.responseText
    End If
    
    Exit Function
    
ErrorHandler:
    CallAI_API = "Error: " & Err.Description
End Function

' Simple JSON parsing to extract content
Function ExtractContentFromJSON(jsonResponse As String) As String
    On Error GoTo ErrorHandler
    
    ' Find content field
    Dim contentStart As Long
    Dim contentEnd As Long
    Dim result As String
    
    contentStart = InStr(jsonResponse, """content"":")
    If contentStart = 0 Then
        ExtractContentFromJSON = "Unable to parse response"
        Exit Function
    End If
    
    ' Locate the start position of content value
    contentStart = InStr(contentStart, jsonResponse, """:")
    If contentStart = 0 Then
        ExtractContentFromJSON = "Unable to parse response"
        Exit Function
    End If
    
    contentStart = contentStart + 3 ' Move past the first quote
    
    ' Locate the end position of content value
    contentEnd = contentStart
    Do While contentEnd < Len(jsonResponse)
        contentEnd = InStr(contentEnd + 1, jsonResponse, """")
        If contentEnd = 0 Then
            ExtractContentFromJSON = "Unable to parse response"
            Exit Function
        End If
        
        ' Check if it's an escaped quote
        If Mid(jsonResponse, contentEnd - 1, 1) <> "\" Then
            Exit Do
        End If
    Loop
    
    If contentEnd = 0 Then
        ExtractContentFromJSON = "Unable to parse response"
        Exit Function
    End If
    
    result = Mid(jsonResponse, contentStart, contentEnd - contentStart)
    ' Handle escape characters
    result = Replace(result, "\""", """")
    result = Replace(result, "\n", vbCrLf)
    result = Replace(result, "\t", vbTab)
    result = Replace(result, "\r", "")
    
    ExtractContentFromJSON = result
    Exit Function
    
ErrorHandler:
    ExtractContentFromJSON = "Parse error: " & Err.Description
End Function

' Escape special characters in JSON strings
Function EscapeJSON(inputStr As String) As String
    inputStr = Replace(inputStr, "\", "\\")
    inputStr = Replace(inputStr, """", "\""")
    inputStr = Replace(inputStr, vbCrLf, "\n")
    inputStr = Replace(inputStr, vbCr, "\r")
    inputStr = Replace(inputStr, vbLf, "\n")
    inputStr = Replace(inputStr, vbTab, "\t")
    EscapeJSON = inputStr
End Function




' Reset AI configuration
Public Sub ResetAIConfig()
    ' Clear global variables
    g_APIKey = ""
    g_Model = ""
    g_APIEndpoint = ""
    
    ' Clear configuration sheet
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("_AIConfig_")
    If Not ws Is Nothing Then
        ws.Range("A1:D10").ClearContents
    End If
    
    ' Prompt user to reconfigure
    MsgBox "Configuration has been reset. Please reconfigure the model settings.", vbInformation, "Reset Successful"
    
    ' Call setup function
    SetupMyAI
End Sub

' Show current configuration info
Private Sub ShowAIConfig()
    MsgBox "Current AI configuration:" & vbCrLf & _
           "API key: " & IIf(g_APIKey = "", "Not configured", "Configured") & vbCrLf & _
           "Model: " & g_Model & vbCrLf & _
           "API endpoint: " & g_APIEndpoint, vbInformation
End Sub

' Setup My AI model with user input and validation
Sub SetupMyAI()
    ' Load saved configuration first
    LoadAIConfig
    
    ' Check if already configured
    If g_APIKey <> "" And g_Model <> "" And g_APIEndpoint <> "" Then
        MsgBox "AI model is already configured!" & vbCrLf & _
               "Model: " & g_Model & vbCrLf & _
               "API Endpoint: " & g_APIEndpoint, vbInformation
        Exit Sub
    End If
    
    Dim modelName As String
    Dim apiEndpoint As String
    Dim apiKey As String
    
    ' Prompt user for configuration
    modelName = InputBox("Enter the model name (e.g., qwen-plus):", "Model Configuration")
    If modelName = "" Then Exit Sub
    
    apiEndpoint = InputBox("Enter the API endpoint:", "API Configuration")
    If apiEndpoint = "" Then Exit Sub
    
    apiKey = InputBox("Enter the API key:", "API Key Configuration")
    If apiKey = "" Then Exit Sub
    
    ' Configure and test connection
    ConfigureAIModel apiKey, modelName, apiEndpoint
    
    ' Save configuration to VBA macro
    SaveAIConfig
End Sub

' Save AI configuration to worksheet
Private Sub SaveAIConfig()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET)
    
    ' Create config sheet if not exists
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = CONFIG_SHEET
        ws.Visible = xlSheetVeryHidden
    End If
    
    ' Store config values
    ws.Range("A1").Value = "APIKey"
    ws.Range("B1").Value = g_APIKey
    ws.Range("A2").Value = "Model"
    ws.Range("B2").Value = g_Model
    ws.Range("A3").Value = "APIEndpoint"
    ws.Range("B3").Value = g_APIEndpoint
    
    ThisWorkbook.Save
End Sub

' Load AI configuration from worksheet
Private Sub LoadAIConfig()
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET)
    
    If Not ws Is Nothing Then
        g_APIKey = ws.Range("B1").Value
        g_Model = ws.Range("B2").Value
        g_APIEndpoint = ws.Range("B3").Value
    End If
End Sub

