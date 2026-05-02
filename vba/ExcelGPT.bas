Attribute VB_Name = "ExcelGPT"
Option Explicit

Private Const API_URL As String = "http://127.0.0.1:8000/ask"
Private Const REQUEST_TIMEOUT_MS As Long = 60000

Public Sub PerguntarGPT()
    Dim sourceCell As Range
    Dim prompt As String
    Dim responseText As String
    Dim wasSuccessful As Boolean
    Dim answer As String
    Dim errorMessage As String

    On Error GoTo ConnectionError

    If TypeName(Selection) <> "Range" Then
        MsgBox "Select a cell before running the macro.", vbExclamation, "ExcelGPT"
        Exit Sub
    End If

    Set sourceCell = ActiveCell

    If IsError(sourceCell.Value) Then
        MsgBox "The active cell contains an Excel error. Enter text and try again.", vbExclamation, "ExcelGPT"
        Exit Sub
    End If

    prompt = Trim$(CStr(sourceCell.Value))

    If Len(prompt) = 0 Then
        MsgBox "The active cell is empty. Type a prompt and try again.", vbInformation, "ExcelGPT"
        Exit Sub
    End If

    Application.StatusBar = "Sending prompt to the local OpenAI API..."
    responseText = SendAskRequest(prompt)

    If Len(responseText) = 0 Then
        MsgBox "The local API returned an empty response.", vbExclamation, "ExcelGPT"
        GoTo CleanExit
    End If

    wasSuccessful = ExtractJsonBoolean(responseText, "success")

    If wasSuccessful Then
        answer = ExtractJsonString(responseText, "answer")

        If Len(answer) = 0 Then
            MsgBox "The API returned success, but no answer was found.", vbExclamation, "ExcelGPT"
            GoTo CleanExit
        End If

        sourceCell.Offset(0, 1).Value = answer
        sourceCell.Offset(0, 1).WrapText = True
    Else
        errorMessage = ExtractJsonString(responseText, "error")

        If Len(errorMessage) = 0 Then
            errorMessage = "The local API returned an unexpected error."
        End If

        MsgBox errorMessage, vbExclamation, "ExcelGPT"
    End If

CleanExit:
    Application.StatusBar = False
    Exit Sub

ConnectionError:
    Application.StatusBar = False
    MsgBox "Could not connect to the local backend." & vbCrLf & vbCrLf & _
           "Make sure FastAPI is running at " & API_URL & ".", vbExclamation, "ExcelGPT"
End Sub

Private Function SendAskRequest(ByVal prompt As String, Optional ByVal systemPrompt As String = "") As String
    Dim http As Object
    Dim requestBody As String

    requestBody = BuildRequestJson(prompt, systemPrompt)

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", API_URL, False
    http.setTimeouts REQUEST_TIMEOUT_MS, REQUEST_TIMEOUT_MS, REQUEST_TIMEOUT_MS, REQUEST_TIMEOUT_MS
    http.setRequestHeader "Content-Type", "application/json"
    http.send requestBody

    If http.Status = 0 Then
        Err.Raise vbObjectError + 1000, "ExcelGPT", "No HTTP status was returned by the local backend."
    End If

    SendAskRequest = CStr(http.responseText)
End Function

Private Function BuildRequestJson(ByVal prompt As String, Optional ByVal systemPrompt As String = "") As String
    Dim json As String

    json = "{""prompt"":""" & EscapeJsonString(prompt) & """"

    If Len(systemPrompt) > 0 Then
        json = json & ",""system_prompt"":""" & EscapeJsonString(systemPrompt) & """"
    End If

    json = json & "}"
    BuildRequestJson = json
End Function

Private Function EscapeJsonString(ByVal value As String) As String
    Dim index As Long
    Dim currentChar As String
    Dim charCode As Long
    Dim result As String

    For index = 1 To Len(value)
        currentChar = Mid$(value, index, 1)

        Select Case currentChar
            Case """"
                result = result & "\"""
            Case "\"
                result = result & "\\"
            Case vbBack
                result = result & "\b"
            Case vbFormFeed
                result = result & "\f"
            Case vbCr
                result = result & "\r"
            Case vbLf
                result = result & "\n"
            Case vbTab
                result = result & "\t"
            Case Else
                charCode = AscW(currentChar)

                If charCode < 32 Then
                    result = result & "\u" & Right$("0000" & Hex$(charCode), 4)
                Else
                    result = result & currentChar
                End If
        End Select
    Next index

    EscapeJsonString = result
End Function

Private Function ExtractJsonBoolean(ByVal jsonText As String, ByVal key As String) As Boolean
    Dim rawValue As String

    rawValue = LCase$(ExtractJsonRawValue(jsonText, key))
    ExtractJsonBoolean = (rawValue = "true")
End Function

Private Function ExtractJsonRawValue(ByVal jsonText As String, ByVal key As String) As String
    Dim valueStart As Long
    Dim currentPosition As Long
    Dim currentChar As String

    valueStart = FindJsonValueStart(jsonText, key)

    If valueStart = 0 Then
        Exit Function
    End If

    currentPosition = SkipWhitespace(jsonText, valueStart)

    Do While currentPosition <= Len(jsonText)
        currentChar = Mid$(jsonText, currentPosition, 1)

        If currentChar = "," Or currentChar = "}" Or currentChar = vbCr Or currentChar = vbLf Then
            Exit Do
        End If

        currentPosition = currentPosition + 1
    Loop

    ExtractJsonRawValue = Trim$(Mid$(jsonText, SkipWhitespace(jsonText, valueStart), currentPosition - SkipWhitespace(jsonText, valueStart)))
End Function

Private Function ExtractJsonString(ByVal jsonText As String, ByVal key As String) As String
    Dim valueStart As Long

    valueStart = FindJsonValueStart(jsonText, key)

    If valueStart = 0 Then
        Exit Function
    End If

    valueStart = SkipWhitespace(jsonText, valueStart)

    If valueStart = 0 Then
        Exit Function
    End If

    If Mid$(jsonText, valueStart, 1) <> """" Then
        Exit Function
    End If

    ExtractJsonString = ReadJsonStringToken(jsonText, valueStart)
End Function

Private Function FindJsonValueStart(ByVal jsonText As String, ByVal key As String) As Long
    Dim keyPattern As String
    Dim keyPosition As Long
    Dim colonPosition As Long

    keyPattern = """" & key & """"
    keyPosition = InStr(1, jsonText, keyPattern, vbTextCompare)

    If keyPosition = 0 Then
        Exit Function
    End If

    colonPosition = InStr(keyPosition + Len(keyPattern), jsonText, ":")

    If colonPosition = 0 Then
        Exit Function
    End If

    FindJsonValueStart = colonPosition + 1
End Function

Private Function ReadJsonStringToken(ByVal jsonText As String, ByVal openingQuotePosition As Long) As String
    Dim index As Long
    Dim currentChar As String
    Dim escapeChar As String
    Dim hexValue As String
    Dim result As String

    index = openingQuotePosition + 1

    Do While index <= Len(jsonText)
        currentChar = Mid$(jsonText, index, 1)

        If currentChar = "\" Then
            index = index + 1

            If index > Len(jsonText) Then
                Exit Do
            End If

            escapeChar = Mid$(jsonText, index, 1)

            Select Case escapeChar
                Case """"
                    result = result & """"
                Case "\"
                    result = result & "\"
                Case "/"
                    result = result & "/"
                Case "b"
                    result = result & vbBack
                Case "f"
                    result = result & vbFormFeed
                Case "n"
                    result = result & vbLf
                Case "r"
                    result = result & vbCr
                Case "t"
                    result = result & vbTab
                Case "u"
                    If index + 4 <= Len(jsonText) Then
                        hexValue = Mid$(jsonText, index + 1, 4)
                        result = result & UnicodeEscapeToString(hexValue)
                        index = index + 4
                    End If
                Case Else
                    result = result & escapeChar
            End Select
        ElseIf currentChar = """" Then
            Exit Do
        Else
            result = result & currentChar
        End If

        index = index + 1
    Loop

    ReadJsonStringToken = result
End Function

Private Function UnicodeEscapeToString(ByVal hexValue As String) As String
    On Error GoTo InvalidHexValue

    UnicodeEscapeToString = ChrW$(CLng("&H" & hexValue))
    Exit Function

InvalidHexValue:
    UnicodeEscapeToString = ""
End Function

Private Function SkipWhitespace(ByVal textValue As String, ByVal startPosition As Long) As Long
    Dim index As Long
    Dim currentChar As String

    If startPosition <= 0 Then
        Exit Function
    End If

    For index = startPosition To Len(textValue)
        currentChar = Mid$(textValue, index, 1)

        If currentChar <> " " And currentChar <> vbTab And currentChar <> vbCr And currentChar <> vbLf Then
            SkipWhitespace = index
            Exit Function
        End If
    Next index
End Function
