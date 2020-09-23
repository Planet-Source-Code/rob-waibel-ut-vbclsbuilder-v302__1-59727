Global Const vbLongDate As Integer = 9

' Date format string for SQL Server
Global Const SQL_DATE_FORMAT As String = "\'yyyymmdd\'"
Global Const SQL_DATETIME_FORMAT As String = "\'yyyymmdd Hh:Nn:Ss.000\'"

Public Function SQLData(ByVal DataVal As Variant, Optional ByVal DataType As Integer) As String

    Dim result As String
    Dim QuotePos As Integer
    
    If DataType = 0 Then
        DataType = vbString
    End If
    
    If IsNull(DataVal) Then
        result = "NULL"
    Else
        Select Case DataType
            Case vbNull
                result = "NULL"
            Case vbString
                result = DataVal & ""
                QuotePos = InStr(1, result, "'")
                Do While QuotePos > 0
                    result = Left(result, QuotePos) & "'" & Mid(result, QuotePos + 1)
                    QuotePos = InStr(QuotePos + 2, result, "'")
                Loop
                If result = "" Then
                    result = "NULL"
                Else
                    result = "'" & Trim(result) & "'"
                End If
            Case vbDate
                result = Format(DataVal, SQL_DATE_FORMAT)
            Case vbBoolean
                result = Abs(CInt(DataVal)) & ""
            Case vbLongDate
                result = Format(DataVal, SQL_DATETIME_FORMAT)
            Case Else
                result = Trim(DataVal & "")
        End Select
    End If
    
    SQLData = result
    
End Function