Attribute VB_Name = "NVDA_API"
Option Explicit

' Place your Alpha Vantage API key here
Private Const API_KEY As String = "VOTRE_CLE_API"

' Fetch the latest NVDA price and place it in cell C3 of the active sheet
Public Sub RecupCoursNVDA()
    Dim url As String
    Dim http As Object
    Dim responseText As String
    Dim price As Double

    url = "https://www.alphavantage.co/query?function=GLOBAL_QUOTE&symbol=NVDA&apikey=" & API_KEY

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.send

    If http.Status = 200 Then
        responseText = http.responseText
        price = ParsePriceFromJson(responseText)

        If price > 0 Then
            ThisWorkbook.ActiveSheet.Range("C3").Value = price
        Else
            MsgBox "Impossible d'extraire le prix dans la réponse JSON.", vbExclamation
        End If
    Else
        MsgBox "Erreur HTTP: " & http.Status & " - " & http.statusText, vbCritical
    End If
End Sub

' Extract the "05. price" field from the Alpha Vantage GLOBAL_QUOTE response
' Handles locales where the decimal separator is a comma by replacing the dot before conversion
Private Function ParsePriceFromJson(jsonText As String) As Double
    Dim key As String
    Dim pos As Long, startPos As Long, endPos As Long
    Dim rawValue As String
    Dim decimalSep As String

    key = """05. price"":"
    pos = InStr(1, jsonText, key, vbTextCompare)

    If pos > 0 Then
        startPos = InStr(pos + Len(key), jsonText, """)
        If startPos > 0 Then
            endPos = InStr(startPos + 1, jsonText, """)
            If endPos > startPos Then
                rawValue = Mid$(jsonText, startPos + 1, endPos - startPos - 1)
                decimalSep = Application.International(xlDecimalSeparator)
                ParsePriceFromJson = CDbl(Replace(rawValue, ".", decimalSep))
                Exit Function
            End If
        End If
    End If

    ParsePriceFromJson = 0
End Function
