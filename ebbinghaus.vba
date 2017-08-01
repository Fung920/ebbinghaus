
Option Explicit

Sub GenerateList()
    Dim xlsSheet As Worksheet
    Dim objRs As Object
    Dim iCount As Integer, iRow As Integer, i As Integer
    Dim dStart As Date, dPreviousDate As Date
    Dim sTitle As String
    Dim arrRule(6) As Integer
    
    Set xlsSheet = Sheet1
    ' iCount = CInt(xlsSheet.Range("B1").Value)
    iCount = 26
    ' dStart = CDate(xlsSheet.Range("D1").Value)
    dStart = CDate("2017/07/26")
    
    arrRule(1) = 1
    arrRule(2) = 2
    arrRule(3) = 4
    arrRule(4) = 7
    arrRule(5) = 15
    arrRule(6) = 30
    
    
    Set objRs = CreateObject("ADODB.Recordset")
    objRs.Fields.Append "Title", 130, 10 'adChar
    objRs.Fields.Append "ReadDate", 7 'aStart
    objRs.Open
    '25A1
    For iRow = 2 To iCount + 1
        objRs.AddNew
        sTitle = ChrW(2 * 16 * 16 * 16 + 5 * 16 * 16 + 160 + 1) & "List" & Format(iRow - 1, "00")
        
        objRs("Title").Value = sTitle
        objRs("ReadDate").Value = dStart
        For i = 1 To 6
            objRs.AddNew
            objRs("Title").Value = sTitle
            objRs("ReadDate").Value = dStart + arrRule(i)
        Next
        
        dStart = dStart + 1
    Next
    objRs.Sort = "ReadDate ASC, Title ASC"
    
    iRow = 2
    
    xlsSheet.Range("A" & iRow).Value = "Date"
    xlsSheet.Range("B" & iRow).Value = "FirstLearn(Before 12:00PM)"
    xlsSheet.Range("C" & iRow).Value = "Review(Before 00:00AM)"
    iRow = iRow + 1
    
    If objRs.RecordCount > 0 Then
        objRs.MoveFirst
        dPreviousDate = objRs("ReadDate").Value
        sTitle = Trim(objRs("Title").Value)
        objRs.MoveNext
    End If
    
    Do Until objRs.EOF
        If dPreviousDate = objRs("ReadDate").Value Then
            dPreviousDate = objRs("ReadDate").Value
            sTitle = sTitle & " ," & Trim(objRs("Title").Value)
        Else
            xlsSheet.Range("A" & iRow).Value = dPreviousDate
            xlsSheet.Range("C" & iRow).Value = sTitle
            
            If iRow <= iCount + 2 Then
                xlsSheet.Range("B" & iRow).Value = ChrW(2 * 16 * 16 * 16 + 5 * 16 * 16 + 160 + 1) & "List" & iRow - 2
            Else
                xlsSheet.Range("B" & iRow).Value = ""
            End If

            iRow = iRow + 1
            
            dPreviousDate = objRs("ReadDate").Value
            sTitle = Trim(objRs("Title").Value)
            
            If iRow Mod 2 = 0 Then SetShadow xlsSheet, iRow
            
        End If
                
        objRs.MoveNext
    Loop
    
    If objRs.RecordCount > 0 Then
        xlsSheet.Range("A" & iRow).Value = dPreviousDate
        xlsSheet.Range("C" & iRow).Value = sTitle
        If iRow <= iCount + 2 Then
            xlsSheet.Range("B" & iRow).Value = "List" & iRow - 2
        Else
            xlsSheet.Range("B" & iRow).Value = ""
        End If
        
        If iRow Mod 2 = 0 Then SetShadow xlsSheet, iRow
    End If
    
    Set objRs = Nothing
    
    MsgBox "Proceed successfully!"
        
End Sub

Sub SetShadow(ByVal xlsSheet As Worksheet, ByVal row As Integer)
    xlsSheet.Activate
    xlsSheet.Rows(row & ":" & row).Select

    With Application.Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
End Sub



