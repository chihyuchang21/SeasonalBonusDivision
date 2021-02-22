Public Sub D()
        
    'Read Environ to get the username
    username = Environ("username")
    YearAndSeason = InputBox("Please Enter Year & Season:" & vbCrLf _
    & "i.e. 2020Q4")
    Workbooks.Open "C:\Users\" & username & "\Desktop\" & YearAndSeason & "季獎金調整清冊"
    outsideFolder = "C:\Users\" & username & "\Desktop\" & "季獎金切檔\"
    bp = YearAndSeason & "季獎金調整清冊-"
    fp = YearAndSeason & "季獎金-"
    
    
    Dim wb As Workbook
    Set wb = Workbooks(YearAndSeason & "季獎金調整清冊")
    
    
    Dim ws As Worksheet
    Set ws = wb.Sheets("貼值")
    
 
    cnt_dept = ws.Range("D" & Rows.Count).End(xlUp).Row - 2 'NOTE
    'Application.DisplayAlerts = False
    
    For a = 1 To cnt_dept

        Dim Func2V, Func1V, PlantV, DeptV, SecV As Variant
        
        With ws
            Func2V = .Range("A" & a + 2).Value
            Func1V = .Range("B" & a + 2).Value
            PlantV = .Range("C" & a + 2).Value
            DeptV = .Range("D" & a + 2).Value
            SecV = .Range("E" & a + 2).Value
            MgV = .Range("F" & a + 2).Value
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
        
        'If MgV = 0 Then
            'For Each Sheet In ActiveWorkbook.Worksheets
                'If Sheet.Name = "主管" Then
                    'Sheet.Delete
                'End If
            'Next
        'End If
        
        Dosomething
        RunCode
        Application.DisplayAlerts = True '1110 added
        ActiveWorkbook.Close SaveChanges:=True
    
    Next
        
End Sub
Sub Dosomething()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    
    For Each xSh In Worksheets
        xSh.Select
        Call RunCode
    Next
    
    Application.ScreenUpdating = True

End Sub
Sub RunCode()
    Dim cnt_ppl As Integer
    Dim WS_Count As Integer
    Dim i As Integer
    
    WS_Count = ActiveWorkbook.Worksheets.Count
    Application.DisplayAlerts = False
    
    For Each Sheet In ActiveWorkbook.Worksheets
        If Sheet.Name = "IDL" Then
            Sheet.Delete
        ElseIf Sheet.Name = "DL" Then
            Sheet.Delete
        End If
    Next


    
        cnt_ppl = ActiveSheet.Range("E" & Rows.Count).End(xlUp).Row - 24


        Rows(22).RowHeight = 44.3
        Columns("V:Z").Delete
        
        If cnt_ppl > 1 Then
            Range("P25").Formula = "=Sum(N25:O25)"
            Range("P25").AutoFill Destination:=Range("P25:P" & 25 + cnt_ppl - 1), Type:=xlFillCopy
        
            Range("Q25").Formula = "=J25+P25"
            Range("Q25").AutoFill Destination:=Range("Q25:Q" & 25 + cnt_ppl - 1), Type:=xlFillCopy
        
            Range("R25").Formula = "=IF(Q25<=K25,(Q25-K25)/K25,IF(Q25>=J25,(Q25-J25)/J25,(Q25-J25)/J25))"
            Range("R25").AutoFill Destination:=Range("R25:R" & 25 + cnt_ppl - 1), Type:=xlFillCopy
        
            Range("T25").Formula = "=J25+P25+S25"
            Range("T25").AutoFill Destination:=Range("T25:T" & 25 + cnt_ppl - 1), Type:=xlFillCopy
            
            '0111added
            Range("J25:J" & 25 + cnt_ppl).Select
            With Selection.Validation
                .Delete
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                Operator:=xlLess, Formula1:="0"
                .IgnoreBlank = True
                .InCellDropdown = True
                    .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = "已設定公式勿修改，請於TQM評比金額欄位或是主管調整欄位輸入金額"
                .IMEMode = xlIMEModeNoControl
                .ShowInput = True
                .ShowError = True
            End With
            
'            '0111IFNEEDED
'            Range("K25:K" & 25 + cnt_ppl).Select
'            With Selection.Validation
'                .Delete
'                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
'                Operator:=xlLess, Formula1:="0"
'                .IgnoreBlank = True
'                .InCellDropdown = True
'                    .InputTitle = ""
'                .ErrorTitle = ""
'                .InputMessage = ""
'                .ErrorMessage = "已設定公式勿修改，請於TQM評比金額欄位或是主管調整欄位輸入金額"
'                .IMEMode = xlIMEModeNoControl
'                .ShowInput = True
'                .ShowError = True
'            End With

            Range("P25:P" & 25 + cnt_ppl).Select
            With Selection.Validation
                .Delete
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                Operator:=xlLess, Formula1:="0"
                .IgnoreBlank = True
                .InCellDropdown = True
                    .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = "已設定公式勿修改，請於TQM評比金額欄位或是主管調整欄位輸入金額"
                .IMEMode = xlIMEModeNoControl
                .ShowInput = True
                .ShowError = True
            End With
            
            Range("Q25:Q" & 25 + cnt_ppl).Select
            With Selection.Validation
                .Delete
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                Operator:=xlLess, Formula1:="0"
                .IgnoreBlank = True
                .InCellDropdown = True
                    .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = "已設定公式勿修改，請於TQM評比金額欄位或是主管調整欄位輸入金額"
                .IMEMode = xlIMEModeNoControl
                .ShowInput = True
                .ShowError = True
            End With
            
            Range("R25:R" & 25 + cnt_ppl).Select
            With Selection.Validation
                .Delete
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                Operator:=xlLess, Formula1:="0"
                .IgnoreBlank = True
                .InCellDropdown = True
                    .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = "已設定公式勿修改，請於TQM評比金額欄位或是主管調整欄位輸入金額"
                .IMEMode = xlIMEModeNoControl
                .ShowInput = True
                .ShowError = True
            End With
            
            Range("T25:T" & 25 + cnt_ppl).Select
            With Selection.Validation
                .Delete
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                Operator:=xlLess, Formula1:="0"
                .IgnoreBlank = True
                .InCellDropdown = True
                    .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = "已設定公式勿修改，請於TQM評比金額欄位或是主管調整欄位輸入金額"
                .IMEMode = xlIMEModeNoControl
                .ShowInput = True
                .ShowError = True
            End With
            

            'TOTAL Formula
            Cells.Find(What:="合計", Lookat:=xlWhole).Activate
        
            ActiveCell.Offset(0, 1).Select '往右一格
            ActiveCell.FormulaR1C1 = "=SUM(R[-" & cnt_ppl & "]C:R[-1]C)"
            
           
            ActiveCell.Select
            With Selection.Validation
                .Delete
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                Operator:=xlLess, Formula1:="0"
                .IgnoreBlank = True
                .InCellDropdown = True
                    .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = "已設定公式勿修改，請於TQM評比金額欄位或是主管調整欄位輸入金額"
                .IMEMode = xlIMEModeNoControl
                .ShowInput = True
                .ShowError = True
            End With
            
        
            ActiveCell.Offset(0, 4).Select
            ActiveCell.FormulaR1C1 = "=SUM(R[-" & cnt_ppl & "]C:R[-1]C)"
            
            
            ActiveCell.Select
            With Selection.Validation
                .Delete
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                Operator:=xlLess, Formula1:="0"
                .IgnoreBlank = True
                .InCellDropdown = True
                    .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = "已設定公式勿修改，請於TQM評比金額欄位或是主管調整欄位輸入金額"
                .IMEMode = xlIMEModeNoControl
                .ShowInput = True
                .ShowError = True
            End With
            
        
            ActiveCell.Offset(0, 1).Select
            ActiveCell.FormulaR1C1 = "=SUM(R[-" & cnt_ppl & "]C:R[-1]C)"
            
            
            ActiveCell.Select
            With Selection.Validation
                .Delete
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                Operator:=xlLess, Formula1:="0"
                .IgnoreBlank = True
                .InCellDropdown = True
                    .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = "已設定公式勿修改，請於TQM評比金額欄位或是主管調整欄位輸入金額"
                .IMEMode = xlIMEModeNoControl
                .ShowInput = True
                .ShowError = True
            End With
            
        
            ActiveCell.Offset(0, 1).Select
            ActiveCell.FormulaR1C1 = "=SUM(R[-" & cnt_ppl & "]C:R[-1]C)"
            
            
            ActiveCell.Select
            With Selection.Validation
                .Delete
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                Operator:=xlLess, Formula1:="0"
                .IgnoreBlank = True
                .InCellDropdown = True
                    .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = "已設定公式勿修改，請於TQM評比金額欄位或是主管調整欄位輸入金額"
                .IMEMode = xlIMEModeNoControl
                .ShowInput = True
                .ShowError = True
            End With
            
        
            ActiveCell.Offset(0, 1).Select
            ActiveCell.FormulaR1C1 = "=SUM(R[-" & cnt_ppl & "]C:R[-1]C)"
            
            
            ActiveCell.Select
            With Selection.Validation
                .Delete
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                Operator:=xlLess, Formula1:="0"
                .IgnoreBlank = True
                .InCellDropdown = True
                    .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = "已設定公式勿修改，請於TQM評比金額欄位或是主管調整欄位輸入金額"
                .IMEMode = xlIMEModeNoControl
                .ShowInput = True
                .ShowError = True
            End With
            

            ActiveCell.Offset(0, 2).Select
            ActiveCell.FormulaR1C1 = "=SUM(R[-" & cnt_ppl & "]C:R[-1]C)"
            
            
            ActiveCell.Select
            With Selection.Validation
                .Delete
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                Operator:=xlLess, Formula1:="0"
                .IgnoreBlank = True
                .InCellDropdown = True
                    .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = "已設定公式勿修改，請於TQM評比金額欄位或是主管調整欄位輸入金額"
                .IMEMode = xlIMEModeNoControl
                .ShowInput = True
                .ShowError = True
            End With
            
        
            ActiveCell.Offset(0, 1).Select
            ActiveCell.FormulaR1C1 = "=SUM(R[-" & cnt_ppl & "]C:R[-1]C)"
            
            
            ActiveCell.Select
            With Selection.Validation
                .Delete
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                Operator:=xlLess, Formula1:="0"
                .IgnoreBlank = True
                .InCellDropdown = True
                    .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = "已設定公式勿修改，請於TQM評比金額欄位或是主管調整欄位輸入金額"
                .IMEMode = xlIMEModeNoControl
                .ShowInput = True
                .ShowError = True
            End With
            
        
        
            Rows("24:24").Select
            Selection.RowHeight = 53.3
                  
            Rows("25:25").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.RowHeight = 30
            Columns("H:H").Select
            Range("A1").Select
            ActiveWindow.Zoom = 60

            Columns("A").ColumnWidth = 3
            Columns("B").ColumnWidth = 9.65
            Columns("C").ColumnWidth = 9.11
            Columns("D").ColumnWidth = 8.33
            Columns("E").ColumnWidth = 8.11
            Columns("F").ColumnWidth = 5.33
            Columns("G").ColumnWidth = 6.22
            Columns("H").ColumnWidth = 13
            Columns("I").ColumnWidth = 10.33
            Columns("J").ColumnWidth = 13.56
            Columns("K").ColumnWidth = 9.89
            Columns("L").ColumnWidth = 5.89
            Columns("M").ColumnWidth = 9.2
            Columns("N").ColumnWidth = 11.56
            Columns("O").ColumnWidth = 9.26
            Columns("P").ColumnWidth = 10.56
            Columns("Q").ColumnWidth = 13.89
            Columns("R").ColumnWidth = 11.89
            Columns("S").ColumnWidth = 12
            Columns("T").ColumnWidth = 13.89
            Columns("U").ColumnWidth = 14
            Columns("V").ColumnWidth = 8.11
            Columns("W").ColumnWidth = 8.11
            Columns("X").ColumnWidth = 8.11
            Columns("Y").ColumnWidth = 8.11
            Columns("Z").ColumnWidth = 8.11
            
            '0111Added
            'ActiveSheet.PageSetup.PrintTitleRows = ActiveSheet.Rows("22:24").Address
            'ActiveSheet.PageSetup.PrintArea = "$A:$U"
        Else
            Range("P25").Formula = "=Sum(N25:O25)"
            Range("Q25").Formula = "=J25+P25"
            Range("R25").Formula = "=IF(Q25<=K25,(Q25-K25)/K25,IF(Q25>=J25,(Q25-J25)/J25,(Q25-J25)/J25))"
            Range("T25").Formula = "=J25+P25+S25"
            
            Range("J25:J26").Select
            With Selection.Validation
                .Delete
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                Operator:=xlLess, Formula1:="0"
                .IgnoreBlank = True
                .InCellDropdown = True
                    .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = "已設定公式勿修改，請於TQM評比金額欄位或是主管調整欄位輸入金額"
                .IMEMode = xlIMEModeNoControl
                .ShowInput = True
                .ShowError = True
            End With
            
'            '0111IFNEEDED
'            Range("K25:K26").Select
'            With Selection.Validation
'                .Delete
'                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
'                Operator:=xlLess, Formula1:="0"
'                .IgnoreBlank = True
'                .InCellDropdown = True
'                    .InputTitle = ""
'                .ErrorTitle = ""
'                .InputMessage = ""
'                .ErrorMessage = "已設定公式勿修改，請於TQM評比金額欄位或是主管調整欄位輸入金額"
'                .IMEMode = xlIMEModeNoControl
'                .ShowInput = True
'                .ShowError = True
'            End With
            

            Range("P25:P26").Select
            With Selection.Validation
                .Delete
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                Operator:=xlLess, Formula1:="0"
                .IgnoreBlank = True
                .InCellDropdown = True
                    .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = "已設定公式勿修改，請於TQM評比金額欄位或是主管調整欄位輸入金額"
                .IMEMode = xlIMEModeNoControl
                .ShowInput = True
                .ShowError = True
            End With
            
            Range("Q25:Q26").Select
            
            With Selection.Validation
                .Delete
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                Operator:=xlLess, Formula1:="0"
                .IgnoreBlank = True
                .InCellDropdown = True
                    .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = "已設定公式勿修改，請於TQM評比金額欄位或是主管調整欄位輸入金額"
                .IMEMode = xlIMEModeNoControl
                .ShowInput = True
                .ShowError = True
            End With
            
            Range("R25:R26").Select
            
            With Selection.Validation
                .Delete
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                Operator:=xlLess, Formula1:="0"
                .IgnoreBlank = True
                .InCellDropdown = True
                    .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = "已設定公式勿修改，請於TQM評比金額欄位或是主管調整欄位輸入金額"
                .IMEMode = xlIMEModeNoControl
                .ShowInput = True
                .ShowError = True
            End With
            
            Range("T25:T26").Select
            
            With Selection.Validation
                .Delete
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                Operator:=xlLess, Formula1:="0"
                .IgnoreBlank = True
                .InCellDropdown = True
                    .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = "已設定公式勿修改，請於TQM評比金額欄位或是主管調整欄位輸入金額"
                .IMEMode = xlIMEModeNoControl
                .ShowInput = True
                .ShowError = True
            End With
            
            Cells.Find(What:="合計", Lookat:=xlWhole).Activate
            ActiveCell.Offset(0, 1).Select '往右一格
            ActiveCell.FormulaR1C1 = "=SUM(R[-" & cnt_ppl & "]C:R[-1]C)"
            
            
            ActiveCell.Select
            With Selection.Validation
                .Delete
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                Operator:=xlLess, Formula1:="0"
                .IgnoreBlank = True
                .InCellDropdown = True
                    .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = "已設定公式勿修改，請於TQM評比金額欄位或是主管調整欄位輸入金額"
                .IMEMode = xlIMEModeNoControl
                .ShowInput = True
                .ShowError = True
            End With
            
        
            ActiveCell.Offset(0, 4).Select
            ActiveCell.FormulaR1C1 = "=SUM(R[-" & cnt_ppl & "]C:R[-1]C)"

            
            ActiveCell.Select
            With Selection.Validation
                .Delete
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                Operator:=xlLess, Formula1:="0"
                .IgnoreBlank = True
                .InCellDropdown = True
                    .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = "已設定公式勿修改，請於TQM評比金額欄位或是主管調整欄位輸入金額"
                .IMEMode = xlIMEModeNoControl
                .ShowInput = True
                .ShowError = True
            End With
            
        
            ActiveCell.Offset(0, 1).Select
            ActiveCell.FormulaR1C1 = "=SUM(R[-" & cnt_ppl & "]C:R[-1]C)"
            
            
            ActiveCell.Select
            With Selection.Validation
                .Delete
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                Operator:=xlLess, Formula1:="0"
                .IgnoreBlank = True
                .InCellDropdown = True
                    .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = "已設定公式勿修改，請於TQM評比金額欄位或是主管調整欄位輸入金額"
                .IMEMode = xlIMEModeNoControl
                .ShowInput = True
                .ShowError = True
            End With
           
        
            ActiveCell.Offset(0, 1).Select
            ActiveCell.FormulaR1C1 = "=SUM(R[-" & cnt_ppl & "]C:R[-1]C)"
            
            
            ActiveCell.Select
            With Selection.Validation
                .Delete
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                Operator:=xlLess, Formula1:="0"
                .IgnoreBlank = True
                .InCellDropdown = True
                    .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = "已設定公式勿修改，請於TQM評比金額欄位或是主管調整欄位輸入金額"
                .IMEMode = xlIMEModeNoControl
                .ShowInput = True
                .ShowError = True
            End With
            
        
            ActiveCell.Offset(0, 1).Select
            ActiveCell.FormulaR1C1 = "=SUM(R[-" & cnt_ppl & "]C:R[-1]C)"
            
           
            ActiveCell.Select
            With Selection.Validation
                .Delete
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                Operator:=xlLess, Formula1:="0"
                .IgnoreBlank = True
                .InCellDropdown = True
                    .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = "已設定公式勿修改，請於TQM評比金額欄位或是主管調整欄位輸入金額"
                .IMEMode = xlIMEModeNoControl
                .ShowInput = True
                .ShowError = True
            End With
            

            ActiveCell.Offset(0, 2).Select
            ActiveCell.FormulaR1C1 = "=SUM(R[-" & cnt_ppl & "]C:R[-1]C)"
            
            
            ActiveCell.Select
            With Selection.Validation
                .Delete
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                Operator:=xlLess, Formula1:="0"
                .IgnoreBlank = True
                .InCellDropdown = True
                    .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = "已設定公式勿修改，請於TQM評比金額欄位或是主管調整欄位輸入金額"
                .IMEMode = xlIMEModeNoControl
                .ShowInput = True
                .ShowError = True
            End With
            
        
            ActiveCell.Offset(0, 1).Select
            ActiveCell.FormulaR1C1 = "=SUM(R[-" & cnt_ppl & "]C:R[-1]C)"
            
            
            ActiveCell.Select
            With Selection.Validation
                .Delete
                .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, _
                Operator:=xlLess, Formula1:="0"
                .IgnoreBlank = True
                .InCellDropdown = True
                    .InputTitle = ""
                .ErrorTitle = ""
                .InputMessage = ""
                .ErrorMessage = "已設定公式勿修改，請於TQM評比金額欄位或是主管調整欄位輸入金額"
                .IMEMode = xlIMEModeNoControl
                .ShowInput = True
                .ShowError = True
            End With
            
            
            Rows("24:24").Select
            Selection.RowHeight = 53.3
                  
            Rows("25:25").Select
            Range(Selection, Selection.End(xlDown)).Select
            Selection.RowHeight = 30
            Columns("H:H").Select
            Range("A1").Select
            ActiveWindow.Zoom = 60

            Columns("A").ColumnWidth = 3
            Columns("B").ColumnWidth = 9.65
            Columns("C").ColumnWidth = 9.11
            Columns("D").ColumnWidth = 8.33
            Columns("E").ColumnWidth = 8.11
            Columns("F").ColumnWidth = 5.33
            Columns("G").ColumnWidth = 6.22
            Columns("H").ColumnWidth = 13
            Columns("I").ColumnWidth = 10.33
            Columns("J").ColumnWidth = 13.56
            Columns("K").ColumnWidth = 9.89
            Columns("L").ColumnWidth = 5.89
            Columns("M").ColumnWidth = 9.2
            Columns("N").ColumnWidth = 11.56
            Columns("O").ColumnWidth = 9.26
            Columns("P").ColumnWidth = 10.56
            Columns("Q").ColumnWidth = 13.89
            Columns("R").ColumnWidth = 11.89
            Columns("S").ColumnWidth = 12
            Columns("T").ColumnWidth = 13.89
            Columns("U").ColumnWidth = 14
            Columns("V").ColumnWidth = 8.11
            Columns("W").ColumnWidth = 8.11
            Columns("X").ColumnWidth = 8.11
            Columns("Y").ColumnWidth = 8.11
            Columns("Z").ColumnWidth = 8.11
            
            '0111Added
            'ActiveSheet.PageSetup.PrintTitleRows = ActiveSheet.Rows("22:24").Address
            'ActiveSheet.PageSetup.PrintArea = "$A:$U"
        End If
        
        

End Sub











