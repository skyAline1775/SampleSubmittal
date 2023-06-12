Attribute VB_Name = "Sample_Submittal"
Private clientCode As String

Sub SampleSubmittal()
'
' SampleSubmittal Macro
'
'

Dim clientArray(0 To 1) As Variant
clientArray(0) = "V3RT"
clientArray(1) = "Other"

With New selectClient
    Call .AddClients(clientArray)
    .Show vbModal
    If .IsClosed Then
        clientCode = .clientCode
    End If
End With

Dim rawBook As Workbook
Dim importBook As Workbook

If StrComp(clientCode, "V3RT") = 0 Then

    For Each sheet In Sheets
    
        sheet.rows(1).delete
        sheet.Columns("D:H").delete
        sheet.Columns(1).delete
        
        For row = 1 To 1000
        
            If isEmpty(sheet.cells(row, 1)) Then Exit For
            
            sheet.cells(row, 1) = Trim(sheet.cells(row, 1) & " " & sheet.cells(row, 2))
        
        Next row
        
        sheet.Columns(2).delete
        sheet.Columns(1).columnWidth = 20
        
    Next

    Exit Sub

End If

If InStr(cells(1, 1), "Drill Sample & QA/QC List") > 0 Then

    For row = 4 To 2000
        If isEmpty(cells(row, 1)) Then Exit For
        
        cells(row, 1) = cells(row, 1) & " " & cells(row, 2) & "-" & cells(row, 3)
        cells(row, 2) = ""
        cells(row, 3) = ""
        
        If InStr(cells(row, 4), "No Sample") > 0 Then cells(row, 1) = cells(row, 1) & " NS"
    Next row
    
    Columns(4).delete
    
    range("1:3").delete

ElseIf StrComp(cells(11, 1), "Drill Hole #") = 0 Then

    Dim holeID As String
    holeID = cells(11, 2)
    
    Set rawBook = ActiveWorkbook
    Set importBook = Workbooks.Add
    
    rawBook.Activate
    cells.Select
    selection.Copy
    
    importBook.Activate
    cells(1, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    For i = 1 To 13
        rows(1).delete
        If StrComp(cells(1, 1), "SampleID") = 0 Then Exit For
    Next i
    
    Columns(1).columnWidth = 25
    
    For row = 1 To 5000

        If isEmpty(cells(row, 1)) And isEmpty(cells(row + 1, 1)) And isEmpty(cells(row + 2, 1)) And isEmpty(cells(row + 3, 1)) Then
            rows(row).delete
            Exit For
        End If
'        If IsEmpty(Cells(row, 2)) Or IsEmpty(Cells(row, 3)) Or StrComp(Cells(row, 1), "SampleID") = 0 Then
        If Not Len(cells(row, 1)) = 10 Or (isEmpty(cells(row, 3)) And Not isEmpty(cells(row, 2))) Then
            rows(row).delete
            row = row - 1
        Else
            If Not isEmpty(cells(row, 2)) Then
                cells(row, 1) = cells(row, 1) & " " & holeID & " " & cells(row, 2) & "-" & cells(row, 3)
            End If
            cells(row, 2) = ""
            cells(row, 3) = ""
            cells(row, 4) = ""
'            If StrComp(Cells(row, 2), "-") = 0 Then Cells(row, 2) = ""
'            Cells(row, 1) = Cells(row, 1) & " " & Cells(row, 2)
'            Cells(row, 2) = ""
'
'            If InStr(Cells(row, 3), "Hypogene") > 0 Then
'                Cells(row, 3) = "Hypogene"
'            Else
'                Cells(row, 3) = ""
'            End If
        End If
        
    Next row

ElseIf StrComp(cells(3, 5), "FMI - Lone Star") = 0 Or InStr(cells(2, 4), "FMI - ") > 0 Then
    
    Set rawBook = ActiveWorkbook
    Set importBook = Workbooks.Add
    
    rawBook.Activate
    cells.Select
    selection.Copy
    
    importBook.Activate
    cells(1, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    For i = 1 To 28
        rows(1).delete
        If StrComp(cells(1, 1), "Sample No") = 0 Then Exit For
    Next i
    
    Columns(1).columnWidth = 25
    
    For row = 1 To 5000

        If isEmpty(cells(row, 1)) And isEmpty(cells(row + 1, 1)) And isEmpty(cells(row + 2, 1)) And isEmpty(cells(row + 3, 1)) Then
            rows(row).delete
            Exit For
        End If
        If isEmpty(cells(row, 1)) Or StrComp(cells(row, 1), "Sample No") = 0 Then
            rows(row).delete
            row = row - 1
        Else
            If StrComp(cells(row, 2), "-") = 0 Then cells(row, 2) = ""
            cells(row, 1) = cells(row, 1) & " " & cells(row, 2)
            cells(row, 2) = ""

            If InStr(cells(row, 3), "Hypogene") > 0 Then
                cells(row, 3) = "Hypogene"
            Else
                cells(row, 3) = ""
            End If
        End If
        
    Next row

ElseIf StrComp(cells(21, 1), "SAMPLE ID") = 0 And StrComp(cells(21, 2), "INTERVAL") = 0 Then

    Set rawBook = ActiveWorkbook
    Set importBook = Workbooks.Add
    
    rawBook.Activate
    rawBook.ActiveSheet.cells.Select
    selection.Copy
    
    importBook.Activate
    cells(1, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    For i = 1 To 21
        rows(1).delete
    Next i
    
    Columns(1).columnWidth = 25
    
    For row = 1 To 5000

        If isEmpty(cells(row, 1)) Then Exit For
        
        cells(row, 1) = cells(row, 1) & " " & cells(row, 2)
        
    Next row
    
    Columns("B:M").delete
    
ElseIf StrComp(cells(1, 1), "SAMPLEID") = 0 And StrComp(cells(1, 2), "Analysis requested") = 0 Then

    Set rawBook = ActiveWorkbook
    Set importBook = Workbooks.Add
    
    rawBook.Activate
    Columns("A:B").Select
    selection.Copy
    
    importBook.Activate
    cells(1, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    rows(1).delete
    
    Columns(1).columnWidth = 25
    Columns(2).columnWidth = 55
    
    Dim lastRow As Integer
    
    For row = 1 To 5000
        If isEmpty(cells(row, 1)) Then
            lastRow = row - 1
            rows(row + 1).delete
            Exit For
        End If
    Next row
    
    Dim counter As Integer
    counter = 1
    
    Dim jobPrefix As String
    jobPrefix = Split(cells(1, 1), "_")(arrayLength(Split(cells(1, 1), "_")) - 1)
    jobPrefix = Replace(cells(1, 1), "_" & jobPrefix, "")
    
    ActiveSheet.name = "Combined"
    
    Dim temp As String
    For row = 2 To lastRow
        temp = Split(cells(row, 1), "_")(arrayLength(Split(cells(row, 1), "_")) - 1)
        temp = Replace(cells(row, 1), "_" & temp, "")
        If Not StrComp(temp, jobPrefix) = 0 Then
            range(cells(counter, 1), cells(row - 1, 2)).Copy
            Sheets.Add after:=Sheets(Sheets.count)
            Sheets(Sheets.count).name = jobPrefix
            cells(1, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Columns(1).columnWidth = 25
            Columns(2).columnWidth = 55
            Sheets(1).Activate
            counter = row
            jobPrefix = temp
        End If
    Next row
    
    range(cells(counter, 1), cells(lastRow, 2)).Copy
    Sheets.Add after:=Sheets(Sheets.count)
    Sheets(Sheets.count).name = jobPrefix
    cells(1, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Columns(1).columnWidth = 25
    Columns(2).columnWidth = 55
    
ElseIf StrComp(cells(1, 1), "Prep Code") = 0 Then

    

End If

End Sub

Private Function arrayLength(arr As Variant) As Integer
    arrayLength = UBound(arr) - LBound(arr) + 1
End Function
