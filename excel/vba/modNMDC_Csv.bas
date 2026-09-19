Attribute VB_Name = "modNMDC_Csv"
Option Explicit

Public Function NMDC_ParseCsvFile(ByVal filePath As String, ByRef headers As Variant, ByRef data As Variant, ByRef dataCount As Long, ByRef columnCount As Long) As Boolean
    On Error GoTo Handler

    Dim text As String
    Dim rowCount As Long
    Dim headerValues() As Variant
    Dim dataValues() As Variant

    text = NMDC_ReadUtf8Text(filePath)
    If Len(text) = 0 Then
        NMDC_ParseCsvFile = False
        Exit Function
    End If

    NMDC_CsvDimensions text, rowCount, columnCount
    If rowCount < 1 Or columnCount < 1 Then
        NMDC_ParseCsvFile = False
        Exit Function
    End If

    ReDim headerValues(1 To 1, 1 To columnCount)
    dataCount = rowCount - 1
    If dataCount > 0 Then
        ReDim dataValues(1 To dataCount, 1 To columnCount)
    Else
        ReDim dataValues(1 To 1, 1 To columnCount)
    End If

    NMDC_FillCsvArrays text, headerValues, dataValues, rowCount, columnCount
    headers = headerValues
    data = dataValues
    NMDC_ParseCsvFile = True
    Exit Function

Handler:
    NMDC_LogError "CSV_PARSE_ERROR", _
        "Excel could not parse an exported CSV file safely.", _
        "File: " & filePath & " | " & Err.Number & " - " & Err.Description
    NMDC_ParseCsvFile = False
End Function

Private Function NMDC_ReadUtf8Text(ByVal filePath As String) As String
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile filePath
    NMDC_ReadUtf8Text = stream.ReadText
    stream.Close
End Function

Private Sub NMDC_CsvDimensions(ByVal text As String, ByRef rowCount As Long, ByRef columnCount As Long)
    Dim i As Long
    Dim currentColumns As Long
    Dim inQuotes As Boolean
    Dim rowTouched As Boolean
    Dim ch As String
    Dim nextCh As String

    rowCount = 0
    columnCount = 1
    currentColumns = 1

    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch = Chr$(34) Then
            If inQuotes And i < Len(text) Then
                nextCh = Mid$(text, i + 1, 1)
                If nextCh = Chr$(34) Then
                    i = i + 1
                    rowTouched = True
                Else
                    inQuotes = False
                End If
            Else
                inQuotes = Not inQuotes
            End If
            rowTouched = True
        ElseIf Not inQuotes Then
            If ch = "," Then
                currentColumns = currentColumns + 1
                rowTouched = True
            ElseIf ch = vbCr Or ch = vbLf Then
                If ch = vbCr And i < Len(text) Then
                    If Mid$(text, i + 1, 1) = vbLf Then i = i + 1
                End If
                rowCount = rowCount + 1
                If currentColumns > columnCount Then columnCount = currentColumns
                currentColumns = 1
                rowTouched = False
            Else
                rowTouched = True
            End If
        Else
            rowTouched = True
        End If
    Next i

    If rowTouched Then
        rowCount = rowCount + 1
        If currentColumns > columnCount Then columnCount = currentColumns
    End If
End Sub

Private Sub NMDC_FillCsvArrays(ByVal text As String, ByRef headers() As Variant, ByRef data() As Variant, ByVal rowCount As Long, ByVal columnCount As Long)
    Dim i As Long
    Dim rowIndex As Long
    Dim columnIndex As Long
    Dim inQuotes As Boolean
    Dim ch As String
    Dim nextCh As String
    Dim fieldText As String

    rowIndex = 1
    columnIndex = 1
    fieldText = ""

    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If inQuotes Then
            If ch = Chr$(34) Then
                If i < Len(text) And Mid$(text, i + 1, 1) = Chr$(34) Then
                    fieldText = fieldText & Chr$(34)
                    i = i + 1
                Else
                    inQuotes = False
                End If
            ElseIf ch = vbCr Or ch = vbLf Then
                If ch = vbCr And i < Len(text) Then
                    If Mid$(text, i + 1, 1) = vbLf Then i = i + 1
                End If
                If Len(fieldText) > 0 And Right$(fieldText, 1) <> " " Then fieldText = fieldText & " "
            Else
                fieldText = fieldText & ch
            End If
        Else
            Select Case ch
                Case Chr$(34)
                    inQuotes = True
                Case ","
                    NMDC_AssignCsvField headers, data, rowIndex, columnIndex, fieldText
                    fieldText = ""
                    columnIndex = columnIndex + 1
                Case vbCr, vbLf
                    If ch = vbCr And i < Len(text) Then
                        If Mid$(text, i + 1, 1) = vbLf Then i = i + 1
                    End If
                    NMDC_AssignCsvField headers, data, rowIndex, columnIndex, fieldText
                    fieldText = ""
                    rowIndex = rowIndex + 1
                    columnIndex = 1
                    If rowIndex > rowCount Then Exit For
                Case Else
                    fieldText = fieldText & ch
            End Select
        End If
    Next i

    If rowIndex <= rowCount And (Len(fieldText) > 0 Or columnIndex > 1) Then
        NMDC_AssignCsvField headers, data, rowIndex, columnIndex, fieldText
    End If

    If Len(CStr(headers(1, 1))) > 0 Then
        If AscW(Left$(CStr(headers(1, 1)), 1)) = &HFEFF Then headers(1, 1) = Mid$(CStr(headers(1, 1)), 2)
    End If
End Sub

Private Sub NMDC_AssignCsvField(ByRef headers() As Variant, ByRef data() As Variant, ByVal rowIndex As Long, ByVal columnIndex As Long, ByVal value As String)
    If columnIndex < 1 Or columnIndex > UBound(headers, 2) Then Exit Sub
    If rowIndex = 1 Then
        headers(1, columnIndex) = value
    ElseIf rowIndex - 1 <= UBound(data, 1) Then
        data(rowIndex - 1, columnIndex) = value
    End If
End Sub
