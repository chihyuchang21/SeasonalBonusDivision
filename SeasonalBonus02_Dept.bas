Sub SeasonalBonus02_Dept()
    
    '先創檔，若檔存在則 +i切部門，最後再由0判斷要不要刪工種主頁
    
    
    Dim a, b, cnt_dept As Integer
    Dim username, myFolder, YearAndSeason, fileName, OriginalName, deptName, foldernameFunc2 As String
    
    YearAndSeason = InputBox("Please Enter Year & Season:" & vbCrLf _
    & "i.e. 2020Q4")
    
    username = Environ("username")
    myFolder = "C:\Users\" & username & "\Desktop\季獎金切檔"
    Workbooks.Open "C:\Users\" & username & "\Desktop\" & YearAndSeason & "季獎金調整清冊"
    OriginalName = "C:\Users\" & username & "\Desktop\" & YearAndSeason & "季獎金調整清冊"
    fileName = ActiveWorkbook.Name
    
    If InStr(fileName, ".") > 0 Then
        fileName = Left(fileName, InStr(fileName, ".") - 1)
    End If
    
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    'Check the repetition, count 工作表A列非空單元格行數, and -1(處) '暫-2
    cnt_dept = Sheets("貼值").Range("D" & Rows.Count).End(xlUp).Row - 2
    
    For a = 1 To cnt_dept
        Workbooks.Open OriginalName
        '暫設處不同則創檔(需再更細緻)
        If Sheets("貼值").Range("D" & 3 + a).Value <> Sheets("貼值").Range("D" & 2 + a).Value Then
            For b = 1 To 3
                
                With Sheets(b)
                    .Range("$A$24:$Z$10000").AutoFilter Field:=23, Criteria1:="=" & Sheets("貼值").Range("A" & 2 + a).Value, _
                        Operator:=xlOr, Criteria2:="="
    
                    .Range("$A$24:$Z$10000").AutoFilter Field:=24, Criteria1:="=" & Sheets("貼值").Range("B" & 2 + a).Value, _
                        Operator:=xlOr, Criteria2:="="
    
                    .Range("$A$24:$Z$10000").AutoFilter Field:=25, Criteria1:="=" & Sheets("貼值").Range("C" & 2 + a).Value, _
                        Operator:=xlOr, Criteria2:="="
    
                    .Range("$A$24:$Z$10000").AutoFilter Field:=26, Criteria1:="=" & Sheets("貼值").Range("D" & 2 + a).Value, _
                        Operator:=xlOr, Criteria2:="="
                End With
   
                With Sheets(b)
                    With .Cells(1, 1).CurrentRegion
                        With .Offset(1, 0).Resize(.Rows.Count - 1, .Columns.Count)
                            If CBool(Application.Subtotal(103, .Columns(1))) Then
                                .Cells.Copy Destination:=.Cells(.Rows.Count + 1, 1)
                            End If
                            .AutoFilter
                            .Cells(1, 1).Resize(.Rows.Count, 1).EntireRow.Delete
                        End With
                    End With
                End With
                
                '這邊不WORK再想想
                Sheets(b).Select
                Rows("25:25").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.RowHeight = 24
                Columns("H:H").Select
                Selection.ColumnWidth = 13.7
                Range("A1").Select
                ActiveWindow.Zoom = 60

                
            Next
            


        foldernameFunc2 = Sheets("貼值").Range("A" & 2 + a).Value
        foldernameFunc1 = Sheets("貼值").Range("B" & 2 + a).Value
        foldernamePlant = Sheets("貼值").Range("C" & 2 + a).Value
        Sheets("貼值").Delete

        
        

        
        If foldernameFunc2 <> foldernameFunc1 And foldernamePlant = 0 Then
        'Save in Func1 folder
        
        ChDir myFolder & "\" & YearAndSeason & "季獎金" & "-" & foldernameFunc2 & "\" & YearAndSeason & "季獎金" & "-" & foldernameFunc1
    
            For k = 1 To 3
                If Sheets(k).Range("Z25") <> "" Then
                    ActiveWorkbook.SaveAs fileName:= _
                        fileName & "-" & Sheets(k).Range("Z25").Value, FileFormat:= _
                        xlOpenXMLWorkbook, CreateBackup:=False
                        Exit For
                        ActiveWindow.Close
                        Workbooks.Open OriginalName
                End If
            Next
        
        
        ElseIf foldernameFunc2 <> foldernameFunc1 And foldernamePlant <> 0 Then
        
        ChDir myFolder & "\" & YearAndSeason & "季獎金" & "-" & foldernameFunc2 & "\" & YearAndSeason & "季獎金" & "-" & foldernameFunc1 & "\" & YearAndSeason & "季獎金調整清冊" & "-" & foldernamePlant
            
            For k = 1 To 3
                If Sheets(k).Range("Z25") <> "" Then
                    ActiveWorkbook.SaveAs fileName:= _
                        fileName & "-" & Sheets(k).Range("Z25").Value, FileFormat:= _
                        xlOpenXMLWorkbook, CreateBackup:=False
                        Exit For
                        ActiveWindow.Close
                        Workbooks.Open OriginalName
                End If
            Next
        
        ElseIf foldernameFunc2 = foldernameFunc1 And foldernamePlant <> 0 Then
        
        ChDir myFolder & "\" & YearAndSeason & "季獎金" & "-" & foldernameFunc2 & "\" & YearAndSeason & "季獎金調整清冊" & "-" & foldernamePlant
        
            For k = 1 To 3
                If Sheets(k).Range("Z25") <> "" Then
                    ActiveWorkbook.SaveAs fileName:= _
                        fileName & "-" & Sheets(k).Range("Z25").Value, FileFormat:= _
                        xlOpenXMLWorkbook, CreateBackup:=False
                        Exit For
                        ActiveWindow.Close
                        Workbooks.Open OriginalName
                End If
            Next
        

        
        Else
        'Save in Func2 folder
        ChDir myFolder & "\" & YearAndSeason & "季獎金" & "-" & foldernameFunc2
        
            If Sheets(1).Range("Z25").Value <> "" Then
        
            ActiveWorkbook.SaveAs fileName:= _
                fileName & "-" & Sheets(1).Range("Z25").Value, FileFormat:= _
                xlOpenXMLWorkbook, CreateBackup:=False
        
            ElseIf Sheets(2).Range("Z25").Value <> "" Then
            
            ActiveWorkbook.SaveAs fileName:= _
                fileName & "-" & Sheets(2).Range("Z25").Value, FileFormat:= _
                xlOpenXMLWorkbook, CreateBackup:=False
        
            Else
        
            ActiveWorkbook.SaveAs fileName:= _
                fileName & "-" & Sheets(3).Range("Z25").Value, FileFormat:= _
                xlOpenXMLWorkbook, CreateBackup:=False
            
            End If

        
        
        Workbooks.Open OriginalName
        
        End If

        End If
    
    Next
    
    '0110Edited
    For Each Workbook In Workbooks
        Workbook.Close
    Next Workbook

End Sub











