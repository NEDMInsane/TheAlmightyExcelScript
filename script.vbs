
Function rowEnd(sheetNum, stRange) As Integer

 Dim lRow As Long
 Dim lCol As Long
 
 lRow = sheetNum.Cells(Rows.Count, stRange).End(xlUp).row
 
 rowEnd = lRow

 End Function
 
 Function Modify_Cell_Date(cellRange, ws)
    Dim fullDate
    Dim cellDate
    
    fullDate = Date
    
    Dim currMon
    Dim currYr
    Dim cellMon
    Dim cellYr
     
    currMon = Month(fullDate)
    currYr = Year(fullDate)
    
    Dim overdue
    'Debug.Print Sheet2.Range(cellRange).Value
    If IsDate(ws.Range(cellRange).Value) Then
        cellDate = ws.Range(cellRange).Value
        cellMon = Month(cellDate)
        cellYr = Year(cellDate)
        If cellMon < currMon And cellYr = currYr Or cellYr < currYr Then
        'OverDue cell Modifier
            ws.Range(cellRange).Interior.Color = RGB(245, 132, 132)
        ElseIf cellMon = currMon And cellYr = currYr Then
            ws.Range(cellRange).Interior.Color = RGB(245, 205, 132)
        Else
            ws.Range(cellRange).Interior.Color = RGB(132, 245, 162)
        End If
    Else
        ws.Range(cellRange).Value = "No Date"
        ws.Range(cellRange).Interior.Color = RGB(245, 205, 132)
    End If
 
 End Function
 
 Function initStatusPg()
    Sheet11.Range("I42").Value = 0
    Sheet11.Range("H42").Value = 0
    Sheet11.Range("G42").Value = 0
    Sheet11.Range("F42").Value = 0
    Sheet11.Range("E42").Value = 0
    Sheet11.Range("D42").Value = 0
    Sheet11.Range("C42").Value = 0
    Sheet11.Range("B42").Value = 0
    
    Sheet11.Range("I43").Value = 0
    Sheet11.Range("H43").Value = 0
    Sheet11.Range("G43").Value = 0
    Sheet11.Range("F43").Value = 0
    Sheet11.Range("E43").Value = 0
    Sheet11.Range("D43").Value = 0
    Sheet11.Range("C43").Value = 0
    Sheet11.Range("B43").Value = 0
    
    Sheet11.Range("J17", "Q17") = 0
    Sheet11.Range("J18", "Q18") = 0
    Sheet11.Range("J19", "Q19") = 0
    Sheet11.Range("J20", "Q20") = 0
    Sheet11.Range("J21", "Q21") = 0
    Sheet11.Range("J22", "Q22") = 0
    Sheet11.Range("J23", "Q23") = 0
    Sheet11.Range("J24", "Q24") = 0
 End Function
 
 
Sub Squadron_Section_Update()
       
    Dim sectdict As Dictionary
    Set sectdict = New Dictionary
    
    Dim rowPos As Dictionary
    Set rowPos = New Dictionary
        
    Dim lRow As Integer
    'Get last row in the "Squadron" Sheet (Sheet2)
    lRow = rowEnd(Sheet2, 7)
        
    Dim name As String
    Dim sect As String
    
    Dim collCC As New Collection
    Dim collCSS As New Collection
    Dim collMTF As New Collection
    Dim collTRR As New Collection
    Dim collTTA As New Collection
    Dim collTTB As New Collection
    Dim collTTC As New Collection
    Dim collTTF As New Collection
        
    Dim collTrngCell As New Collection
    collTrngCell.Add "D"
    collTrngCell.Add "F"
    collTrngCell.Add "H"
    collTrngCell.Add "J"
    collTrngCell.Add "L"
    collTrngCell.Add "N"
    collTrngCell.Add "P"
    collTrngCell.Add "X"
    
    For i = 7 To lRow
    'Getting each name and what section that name belongs to and adding them to a dictionary
    'Also sets a row postion in the "Squadron" sheet
        name = Sheet2.Cells(i, 1)
        sect = Sheet2.Cells(i, 2)
        sectdict.Add name, sect
        rowPos.Add name, i
        For j = 1 To collTrngCell.Count
            Modify_Cell_Date collTrngCell(j) + CStr(i), Sheet2
        Next j
        
    Next i
    
    'adds each person to the correct section Dictionary so we can add the to the right sheet
    For Each k In sectdict.Keys
        If sectdict(k) = "CC" Or sectdict(k) = "CCF" Or sectdict(k) = "CCS" Or sectdict(k) = "CEM" Then
            'Debug.Print "CC"
            collCC.Add k
        ElseIf sectdict(k) = "CSS" Then
            'Debug.Print "CSS"
            collCSS.Add k
        ElseIf sectdict(k) = "MTF" Then
            'Debug.Print "MTF"
            collMTF.Add k
        ElseIf sectdict(k) = "TRR" Then
            'Debug.Print "TRR"
            collTRR.Add k
        ElseIf sectdict(k) = "TTA" Or sectdict(k) = "TTAB" Or sectdict(k) = "TTAP" Then
            'Debug.Print "TTA"
            collTTA.Add k
        ElseIf sectdict(k) = "TTB" Or sectdict(k) = "TTBP" Or sectdict(k) = "TTBS" Or sectdict(k) = "TTBS (82MDSS)" Then
            'Debug.Print "TTB"
            collTTB.Add k
        ElseIf sectdict(k) = "TTCA" Or sectdict(k) = "TTCB" Or sectdict(k) = "TTCC" Or sectdict(k) = "TTC" Then
            'Debug.Print "TTC"
            collTTC.Add k
        ElseIf sectdict(k) = "TTF" Or sectdict(k) = "TTFB" Or sectdict(k) = "TTFC" Or sectdict(k) = "TTFF" Then
            'Debug.Print "TTF"
            collTTF.Add k
        Else
            Debug.Print k; "Unknown Section"
        End If
    Next k
    'Debug.Print collCC.Count; collCSS.Count; collMTF.Count; collTRR.Count; collTTA.Count; collTTB.Count; collTTC.Count; collTTF.Count
    
    'Adding the names to the correct sheet for their section. Each sheet should be cleared before running this, duplicate
    'or invalid entries might make it through. Still working on this.
    'References the "Squadron" Sheet that way anything we look at elsewhere is coming directly from there.
    
    'The main training we look at is Cyber Awareness(D Col), Force Protection(F Col), SAPR(H Col), CUI(J Col),
    'No FEAR(L Col), Religious Freedom(N Col), OPSEC(P Col), Law of War(X Col)
    For i = 1 To collCC.Count
        Sheet3.Cells(i + 1, 1).Value = "=Squadron!A" + CStr(rowPos(collCC.Item(i)))
        For j = 1 To collTrngCell.Count
            Sheet3.Cells(i + 1, j + 1).Value = "=Squadron!" + collTrngCell(j) + CStr(rowPos(collCC.Item(i)))
        Next j
    Next i
    For i = 1 To collCSS.Count
        Sheet4.Cells(i + 1, 1).Value = "=Squadron!A" + CStr(rowPos(collCSS.Item(i)))
        For j = 1 To collTrngCell.Count
            Sheet4.Cells(i + 1, j + 1).Value = "=Squadron!" + collTrngCell(j) + CStr(rowPos(collCSS.Item(i)))
        Next j
    Next i
    For i = 1 To collMTF.Count
        Sheet5.Cells(i + 1, 1).Value = "=Squadron!A" + CStr(rowPos(collMTF.Item(i)))
        For j = 1 To collTrngCell.Count
            Sheet5.Cells(i + 1, j + 1).Value = "=Squadron!" + collTrngCell(j) + CStr(rowPos(collMTF.Item(i)))
        Next j
    Next i
    For i = 1 To collTRR.Count
        Sheet6.Cells(i + 1, 1).Value = "=Squadron!A" + CStr(rowPos(collTRR.Item(i)))
        For j = 1 To collTrngCell.Count
            Sheet6.Cells(i + 1, j + 1).Value = "=Squadron!" + collTrngCell(j) + CStr(rowPos(collTRR.Item(i)))
        Next j
    Next i
    For i = 1 To collTTA.Count
        Sheet7.Cells(i + 1, 1).Value = "=Squadron!A" + CStr(rowPos(collTTA.Item(i)))
        For j = 1 To collTrngCell.Count
            Sheet7.Cells(i + 1, j + 1).Value = "=Squadron!" + collTrngCell(j) + CStr(rowPos(collTTA.Item(i)))
        Next j
    Next i
    For i = 1 To collTTB.Count
        Sheet8.Cells(i + 1, 1).Value = "=Squadron!A" + CStr(rowPos(collTTB.Item(i)))
        For j = 1 To collTrngCell.Count
            Sheet8.Cells(i + 1, j + 1).Value = "=Squadron!" + collTrngCell(j) + CStr(rowPos(collTTB.Item(i)))
        Next j
    Next i
    For i = 1 To collTTC.Count
        Sheet9.Cells(i + 1, 1).Value = "=Squadron!A" + CStr(rowPos(collTTC.Item(i)))
        For j = 1 To collTrngCell.Count
            Sheet9.Cells(i + 1, j + 1).Value = "=Squadron!" + collTrngCell(j) + CStr(rowPos(collTTC.Item(i)))
        Next j
    Next i
    For i = 1 To collTTF.Count
        Sheet10.Cells(i + 1, 1).Value = "=Squadron!A" + CStr(rowPos(collTTF.Item(i)))
        For j = 1 To collTrngCell.Count
            Sheet10.Cells(i + 1, j + 1).Value = "=Squadron!" + collTrngCell(j) + CStr(rowPos(collTTF.Item(i)))
        Next j
    Next i
       
End Sub

Sub Clear_Sq_Pop()

    Dim lRow As Integer
    'Get last row in the "Squadron" Sheet (Sheet2)
    lRow = rowEnd(Sheet2, 1)
    
    For i = 1 To lRow
        Sheet2.Range("A" + CStr(i), "AC" + CStr(i)).UnMerge
        Sheet2.Range("A" + CStr(i), "AC" + CStr(i)).ClearContents
        Sheet2.Range("A" + CStr(i), "AC" + CStr(i)).Interior.Color = RGB(255, 255, 255)
        Sheet2.Range("A" + CStr(i), "AC" + CStr(i)).ClearFormats
    Next i
    Sheet2.Range("A1").Value = "Paste Ancillary Training here"
    
End Sub

Sub Clear_Sect_Pop()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Range("A1").Value = "Name" Then
            Dim lRow As Integer
            lRow = rowEnd(ws, 1)
            
            For i = 2 To lRow
                ws.Range("A" + CStr(i), "AC" + CStr(i)).UnMerge
                ws.Range("A" + CStr(i), "AC" + CStr(i)).ClearContents
                ws.Range("A" + CStr(i), "AC" + CStr(i)).Interior.Color = RGB(255, 255, 255)
                ws.Range("A" + CStr(i), "AC" + CStr(i)).ClearFormats
            Next i
        End If
    Next ws
End Sub

Sub Highlight_Dates()
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Range("A1").Value = "Name" Then
            Dim lRow As Integer
            lRow = rowEnd(ws, 2)
            
            For i = 2 To lRow
                Modify_Cell_Date "B" + CStr(i), ws
                Modify_Cell_Date "C" + CStr(i), ws
                Modify_Cell_Date "D" + CStr(i), ws
                Modify_Cell_Date "E" + CStr(i), ws
                Modify_Cell_Date "F" + CStr(i), ws
                Modify_Cell_Date "G" + CStr(i), ws
                Modify_Cell_Date "H" + CStr(i), ws
                Modify_Cell_Date "I" + CStr(i), ws
            Next i

        End If
    Next ws
End Sub

Function updateSectionNumbers()
    'Sub calculates all people in section
    Dim lRow As Integer
    'SQ Total
    lRow = rowEnd(Sheet2, 7)
    Sheet11.Range("A42").Value = lRow - 6
    'CC
    lRow = rowEnd(Sheet3, 2)
    Sheet11.Range("B42").Value = lRow - 1
    'CSS
    lRow = rowEnd(Sheet4, 2)
    Sheet11.Range("C42").Value = lRow - 1
    'MTF
    lRow = rowEnd(Sheet5, 2)
    Sheet11.Range("D42").Value = lRow - 1
    'TRR
    lRow = rowEnd(Sheet6, 2)
    Sheet11.Range("E42").Value = lRow - 1
    'TTA
    lRow = rowEnd(Sheet7, 2)
    Sheet11.Range("F42").Value = lRow - 1
    'TTB
    lRow = rowEnd(Sheet8, 2)
    Sheet11.Range("G42").Value = lRow - 1
    'TTC
    lRow = rowEnd(Sheet9, 2)
    Sheet11.Range("H42").Value = lRow - 1
    'TTF
    lRow = rowEnd(Sheet10, 2)
    Sheet11.Range("I42").Value = lRow - 1
  
End Function

 Sub Update_Status_Page()
    Dim ws As Worksheet
        
    Dim overdue As Integer
    overdue = 0
    Dim tngDict As Dictionary
    Set tngDict = New Dictionary
    'Numbered training that way the program could iterate through
    'the entries and cells later in the program
    tngDict.Add 10, 0 'Cyber Awareness
    tngDict.Add 11, 0 'Force Protection
    tngDict.Add 12, 0 'SAPR/Suicide Prevention
    tngDict.Add 13, 0 'CUI
    tngDict.Add 14, 0 'No FEAR
    tngDict.Add 15, 0 'Religious Freedom
    tngDict.Add 16, 0 'OPSEC
    tngDict.Add 17, 0 'Law of War
    
    Dim currMon
    Dim currYr
    currMon = Month(Date)
    currYr = Year(Date)
    Dim cellMon
    Dim cellYr
    
    Call initStatusPg 'Calls Function to initialize certain Cells
    Call updateSectionNumbers 'Calls Function to check the amount of people
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Range("A1").Value = "Name" Then
            Dim lRow As Integer
            lRow = rowEnd(ws, 2)
            overdue = 0
            tngDict(10) = 0
            tngDict(11) = 0
            tngDict(12) = 0
            tngDict(13) = 0
            tngDict(14) = 0
            tngDict(15) = 0
            tngDict(16) = 0
            tngDict(17) = 0
            
            For i = 2 To lRow
                For j = 2 To 9
                    If IsDate(ws.Cells(i, j).Value) Then
                        cellMon = Month(ws.Cells(i, j).Value)
                        cellYr = Year(ws.Cells(i, j).Value)
                        If cellMon < currMon And cellYr = currYr Or cellYr < currYr Then
                            overdue = overdue + 1
                            If j = 2 Then
                                tngDict(10) = tngDict(10) + 1
                            ElseIf j = 3 Then
                                tngDict(11) = tngDict(11) + 1
                            ElseIf j = 4 Then
                                tngDict(12) = tngDict(12) + 1
                            ElseIf j = 5 Then
                                tngDict(13) = tngDict(13) + 1
                            ElseIf j = 6 Then
                                tngDict(14) = tngDict(14) + 1
                            ElseIf j = 7 Then
                                tngDict(15) = tngDict(15) + 1
                            ElseIf j = 8 Then
                                tngDict(16) = tngDict(16) + 1
                            ElseIf j = 9 Then
                                tngDict(17) = tngDict(17) + 1
                            Else
                                Debug.Print "Something went wrong with the TNG Updater"
                            End If
                        End If
                    End If
                Next j
            Next i
        End If

        If overdue > 0 Then
            If ws.name = "CC" Then
                For i = 10 To 17
                    Sheet11.Cells(17, i).Value = tngDict(i)
                Next i
                Sheet11.Range("B43").Value = CStr(overdue)
            ElseIf ws.name = "CSS" Then
                For i = 10 To 17
                    Sheet11.Cells(18, i).Value = tngDict(i)
                Next i
                Sheet11.Range("C43").Value = CStr(overdue)
            ElseIf ws.name = "MTF" Then
                For i = 10 To 17
                    Sheet11.Cells(19, i).Value = tngDict(i)
                Next i
                Sheet11.Range("D43").Value = CStr(overdue)
            ElseIf ws.name = "TRR" Then
                For i = 10 To 17
                    Sheet11.Cells(20, i).Value = tngDict(i)
                Next i
                Sheet11.Range("E43").Value = CStr(overdue)
            ElseIf ws.name = "TTA" Then
                For i = 10 To 17
                    Sheet11.Cells(21, i).Value = tngDict(i)
                Next i
                Sheet11.Range("F43").Value = CStr(overdue)
            ElseIf ws.name = "TTB" Then
                For i = 10 To 17
                    Sheet11.Cells(22, i).Value = tngDict(i)
                Next i
                Sheet11.Range("G43").Value = CStr(overdue)
            ElseIf ws.name = "TTC" Then
                For i = 10 To 17
                    Sheet11.Cells(23, i).Value = tngDict(i)
                Next i
                Sheet11.Range("H43").Value = CStr(overdue)
            ElseIf ws.name = "TTF" Then
                For i = 10 To 17
                    Sheet11.Cells(24, i).Value = tngDict(i)
                Next i
                Sheet11.Range("I43").Value = CStr(overdue)
            Else
                Debug.Print "Something went wrong with overdue script."
            End If
        End If
    Next ws
 End Sub
 
 Sub Clear_All()
    Call Clear_Sq_Pop
    Call Clear_Sect_Pop
 End Sub
 
 Sub Update_all()
    Call Squadron_Section_Update
    Call Highlight_Dates
    Call Update_Status_Page
 End Sub
