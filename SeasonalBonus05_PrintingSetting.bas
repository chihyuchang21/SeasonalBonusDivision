Public Sub p()
        Dosomething
        RunCode
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
    Columns("M").ColumnWidth = 9.2
    Columns("U").ColumnWidth = 14
    ActiveSheet.PageSetup.PrintTitleRows = ActiveSheet.Rows("22:24").Address
    ActiveSheet.PageSetup.PrintArea = "$A:$U"
    Columns("J").ColumnWidth = 13.56
    Columns("B").ColumnWidth = 9.65
    Rows(16).RowHeight = 19.5
    Rows(17).RowHeight = 19.5
    ActiveWindow.Zoom = 80
End Sub
