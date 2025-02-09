Attribute VB_Name = "ExcelLanguageTranslator"

Sub TranslateEntireWorkbook()
    Dim ws As Worksheet
    Dim cell As Range
    Dim usedRange As Range
    Dim originalText As String
    Dim translatedText As String
    
    ' Speed up execution
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Sheets
        On Error Resume Next
        Set usedRange = ws.UsedRange ' Identify all used cells
        On Error GoTo 0

        If Not usedRange Is Nothing Then
            ' Loop through each cell in the used range
            For Each cell In usedRange
                If Not IsEmpty(cell.Value) Then ' Process only non-empty cells
                    originalText = cell.Value
                    
                    ' Call translation function
                    translatedText = GoogleTranslate(originalText, "ja", "en")
                    
                    ' Update cell with translated text (if successful)
                    If translatedText <> "" Then cell.Value = translatedText
                End If
            Next cell
        End If
    Next ws

    ' Restore Excel settings
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "Translation Completed!", vbInformation
End Sub

Function GoogleTranslate(text As String, sourceLang As String, targetLang As String) As String
    Dim httpRequest As Object
    Dim apiUrl As String
    Dim jsonResponse As String
    Dim jsonParser As Object
    Dim translatedText As String
    Dim sentence As Variant

    ' Encode text for URL
    text = WorksheetFunction.EncodeURL(text)
    
    ' Construct Google Translate API URL
    apiUrl = "https://translate.googleapis.com/translate_a/single?client=gtx&sl=" & sourceLang & "&tl=" & targetLang & "&dt=t&q=" & text

    ' Create HTTP request
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    With httpRequest
        .Open "GET", apiUrl, False
        .Send
        jsonResponse = .responseText
    End With

    ' Parse JSON response
    Set jsonParser = JsonConverter.ParseJson(jsonResponse)

    ' Extract translated text
    translatedText = ""
    For Each sentence In jsonParser(1)
        translatedText = translatedText & sentence(1) & " "
    Next sentence

    ' Return final translated text
    GoogleTranslate = Trim(translatedText)

    ' Cleanup
    Set httpRequest = Nothing
    Set jsonParser = Nothing
End Function
