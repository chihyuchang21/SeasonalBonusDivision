Function IsWorkBookOpen(Name As String) As Boolean
    Dim xwb As Workbook
    On Error Resume Next
    Set xwb = Application.Workbooks.Item(Name)
    IsWorkBookOpen = (Not xwb Is Nothing)
End Function
Public Sub AdjustWidth()
                Rows("24:24").Select
                Selection.RowHeight = 53.3
                
                
                Rows("25:25").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.RowHeight = 24.9
                Columns("H:H").Select
                Range("A1").Select
                ActiveWindow.Zoom = 60
                
                Columns("A").ColumnWidth = 3
                Columns("B").ColumnWidth = 11.65
                Columns("C").ColumnWidth = 9.11
                Columns("D").ColumnWidth = 8.33
                Columns("E").ColumnWidth = 8.11
                Columns("F").ColumnWidth = 5.33
                Columns("G").ColumnWidth = 6.22
                Columns("H").ColumnWidth = 13
                Columns("I").ColumnWidth = 10.33
                Columns("J").ColumnWidth = 15.56
                Columns("K").ColumnWidth = 9.89
                Columns("L").ColumnWidth = 5.89
                Columns("M").ColumnWidth = 9
                Columns("N").ColumnWidth = 11.56
                Columns("O").ColumnWidth = 9.56
                Columns("P").ColumnWidth = 10.56
                Columns("Q").ColumnWidth = 13.89
                Columns("R").ColumnWidth = 11.89
                Columns("S").ColumnWidth = 12
                Columns("T").ColumnWidth = 13.89
                Columns("U").ColumnWidth = 14.33
                Columns("V").ColumnWidth = 8.11
                Columns("W").ColumnWidth = 8.11
                Columns("X").ColumnWidth = 8.11
                Columns("Y").ColumnWidth = 8.11
                Columns("Z").ColumnWidth = 8.11
End Sub

Sub SeasonalBonus03_SecSheets()
    
    'Read Environ to get the username
    username = Environ("username")
    YearAndSeason = InputBox("Please Enter Year & Season:" & vbCrLf _
    & "i.e. 2020Q4")
    outsideFolder = "C:\Users\" & username & "\Desktop\" & "季獎金切檔\"
    bp = YearAndSeason & "季獎金調整清冊-"
    fp = YearAndSeason & "季獎金-"
    

    Workbooks.Open "C:\Users\" & username & "\Desktop\" & YearAndSeason & "季獎金調整清冊"

    'Execute the code cnt_dept times
    '???!!!-2 or -1!!!???
    
    Dim wb As Workbook
    Set wb = Workbooks(YearAndSeason & "季獎金調整清冊")
    
    
    Dim ws As Worksheet
    Set ws = wb.Sheets("貼值")
    
 
    cnt_dept = ws.Range("D" & Rows.Count).End(xlUp).Row - 2

    
    For a = 1 To cnt_dept

        Dim Func2V, Func1V, PlantV, DeptV, SecV As Variant
        
        With ws
            Func2V = .Range("A" & a + 2).Value
            Func1V = .Range("B" & a + 2).Value
            PlantV = .Range("C" & a + 2).Value
            DeptV = .Range("D" & a + 2).Value
            SecV = .Range("E" & a + 2).Value
            IDLV = .Range("G" & a + 2).Value
            DLV = .Range("H" & a + 2).Value
        End With
            
            
                
        If Func1V = Func2V And PlantV = 0 Then
            Workbooks.Open outsideFolder & fp & Func2V & "\" & bp & DeptV
                
        ElseIf Func1V = Func2V And PlantV <> 0 Then
            Workbooks.Open outsideFolder & fp & Func2V & "\" & bp & PlantV & "\" & bp & DeptV
                
        ElseIf Func1V <> Func2V And PlantV = 0 Then
            Workbooks.Open outsideFolder & fp & Func2V & "\" & fp & Func1V & "\" & bp & DeptV
                    
        ElseIf Func1V <> Func2V And PlantV <> 0 Then
            Workbooks.Open outsideFolder & fp & Func2V & "\" & fp & Func1V & "\" & bp & PlantV & "\" & bp & DeptV
                
        End If
            
            
            If IDLV = 0 Then

            Else
                Sheets("IDL").Copy After:=Sheets("IDL")
                ActiveSheet.Name = "IDL" & "-" & SecV
                
                With ActiveSheet
                    .Range("$A$24:$Z$10000").AutoFilter Field:=2, Criteria1:="=" & SecV, _
                        Operator:=xlOr, Criteria2:="="
                End With
            
                With ActiveSheet
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
                
                Rows("24:24").Select
                Selection.RowHeight = 53.3
                  
                Rows("25:25").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.RowHeight = 24.9
                Columns("H:H").Select
                Range("A1").Select
                ActiveWindow.Zoom = 60
                
                Columns("A").ColumnWidth = 3
                Columns("B").ColumnWidth = 11.65
                Columns("C").ColumnWidth = 9.11
                Columns("D").ColumnWidth = 8.33
                Columns("E").ColumnWidth = 8.11
                Columns("F").ColumnWidth = 5.33
                Columns("G").ColumnWidth = 6.22
                Columns("H").ColumnWidth = 13
                Columns("I").ColumnWidth = 10.33
                Columns("J").ColumnWidth = 15.56
                Columns("K").ColumnWidth = 9.89
                Columns("L").ColumnWidth = 5.89
                Columns("M").ColumnWidth = 9
                Columns("N").ColumnWidth = 11.56
                Columns("O").ColumnWidth = 9.56
                Columns("P").ColumnWidth = 10.56
                Columns("Q").ColumnWidth = 13.89
                Columns("R").ColumnWidth = 11.89
                Columns("S").ColumnWidth = 12
                Columns("T").ColumnWidth = 13.89
                Columns("U").ColumnWidth = 14.33
                Columns("V").ColumnWidth = 8.11
                Columns("W").ColumnWidth = 8.11
                Columns("X").ColumnWidth = 8.11
                Columns("Y").ColumnWidth = 8.11
                Columns("Z").ColumnWidth = 8.11

                
            End If
            
            
            If DLV = 0 Then
            
            Else
                Sheets("DL").Copy After:=Sheets("DL")
                ActiveSheet.Name = "DL" & "-" & SecV
                
                With ActiveSheet
                    .Range("$A$24:$Z$10000").AutoFilter Field:=2, Criteria1:="=" & SecV, _
                        Operator:=xlOr, Criteria2:="="
                End With
            
                With ActiveSheet
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
                
                
                Rows("24:24").Select
                Selection.RowHeight = 53.3
                  
                Rows("25:25").Select
                Range(Selection, Selection.End(xlDown)).Select
                Selection.RowHeight = 24.9
                Columns("H:H").Select
                Range("A1").Select
                ActiveWindow.Zoom = 60
                
                Columns("A").ColumnWidth = 3
                Columns("B").ColumnWidth = 11.65
                Columns("C").ColumnWidth = 9.11
                Columns("D").ColumnWidth = 8.33
                Columns("E").ColumnWidth = 8.11
                Columns("F").ColumnWidth = 5.33
                Columns("G").ColumnWidth = 6.22
                Columns("H").ColumnWidth = 13
                Columns("I").ColumnWidth = 10.33
                Columns("J").ColumnWidth = 15.56
                Columns("K").ColumnWidth = 9.89
                Columns("L").ColumnWidth = 5.89
                Columns("M").ColumnWidth = 9
                Columns("N").ColumnWidth = 11.56
                Columns("O").ColumnWidth = 9.56
                Columns("P").ColumnWidth = 10.56
                Columns("Q").ColumnWidth = 13.89
                Columns("R").ColumnWidth = 11.89
                Columns("S").ColumnWidth = 12
                Columns("T").ColumnWidth = 13.89
                Columns("U").ColumnWidth = 14.33
                Columns("V").ColumnWidth = 8.11
                Columns("W").ColumnWidth = 8.11
                Columns("X").ColumnWidth = 8.11
                Columns("Y").ColumnWidth = 8.11
                Columns("Z").ColumnWidth = 8.11

            
            End If
            
            ActiveWorkbook.Close SaveChanges:=True

        
    Next
            





'ColumnWidthAdjust



End Sub







