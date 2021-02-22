Sub SeasonalBonus01_Directory()
    
    Dim username, YearAndSeason As String
    Dim func2_folder, func1_folder, plant_folder$
    Dim a, cnt_colA As Integer

    
    'Read the username of the computer, and create "季獎金切檔" folder in Desktop.
    
    username = Environ("username")
    myFolder = "C:\Users\" & username & "\Desktop\季獎金切檔"
    
    
    'If the folder does not exist then create, if so, put the files in it.
    
    If Dir(myFolder, vbDirectory) = "" Then
        MkDir myFolder
    
    End If
    
    YearAndSeason = InputBox("Please Enter Year & Season:" & vbCrLf _
    & "i.e. 2020Q4")
    
    Workbooks.Open "C:\Users\" & username & "\Desktop\" & YearAndSeason & "季獎金調整清冊"
    
    'Check the repetition, count 工作表A列非空單元格行數, and -3(計數、功能二)
    cnt_colA = Sheets("貼值").Range("A" & Rows.Count).End(xlUp).Row - 2
    
    For a = 1 To cnt_colA
        
        'O = same ; X = not the same
        Sheets("貼值").Select
        'Func2 X (功二轉換)
        If Range("A" & 3 + a).Value <> Range("A" & 2 + a).Value Then
                
                func2_folder = "C:\Users\" & username & "\Desktop\季獎金切檔\" & YearAndSeason & "季獎金" & "-" & Range("A" & 2 + a)
                
                If Dir(func2_folder, vbDirectory) = "" Then
                    MkDir func2_folder
                End If
            
            'Func1 <> Func2
            If Range("B" & 2 + a).Value <> Range("A" & 2 + a).Value Then
                func1_folder = func2_folder & "\" & YearAndSeason & "季獎金" & "-" & Range("B" & 2 + a)
                
                'If func1_folder DNE
                If Dir(func1_folder, vbDirectory) = "" Then
                    MkDir func1_folder
                ElseIf Range("C" & 2 + a).Value <> 0 Then
                    plant_folder = func1_folder & "\" & YearAndSeason & "季獎金調整清冊" & "-" & Range("C" & 2 + a)
                    'If plant_folder DNE
                    If Dir(plant_folder, vbDirectory) = "" Then
                        MkDir plant_folder
                    End If
                End If
                
                
            'Func1 = Func2
            End If
    
        'Func2 O (功二相同)
        Else
            If Range("B" & 2 + a).Value <> Range("A" & 2 + a).Value Then
                
                func2_folder = "C:\Users\" & username & "\Desktop\季獎金切檔\" & YearAndSeason & "季獎金" & "-" & Range("A" & 2 + a)
                
                'If func2_folder DNE
                If Dir(func2_folder, vbDirectory) = "" Then
                    MkDir func2_folder
                End If
                
                
                func1_folder = func2_folder & "\" & YearAndSeason & "季獎金" & "-" & Range("B" & 2 + a)
                
                'If func1_folder DNE
                If Dir(func1_folder, vbDirectory) = "" Then
                    MkDir func1_folder
                ElseIf Range("C" & 2 + a).Value <> 0 Then
                    plant_folder = func1_folder & "\" & YearAndSeason & "季獎金調整清冊" & "-" & Range("C" & 2 + a)
                    'If plant_folder DNE
                    If Dir(plant_folder, vbDirectory) = "" Then
                        MkDir plant_folder
                    End If
                End If
                        
            '0110Edited:func1==func2 & plant exists
            ElseIf Range("B" & 2 + a).Value = Range("A" & 2 + a).Value Then
                If Range("C" & 2 + a).Value <> 0 Then
                    func2_folder = "C:\Users\" & username & "\Desktop\季獎金切檔\" & YearAndSeason & "季獎金" & "-" & Range("A" & 2 + a)
                    
                    If Dir(func2_folder, vbDirectory) = "" Then
                        MkDir func2_folder
                    End If
                    
                    plant_folder = func2_folder & "\" & YearAndSeason & "季獎金調整清冊" & "-" & Range("C" & 2 + a)
                    If Dir(plant_folder, vbDirectory) = "" Then
                        MkDir plant_folder
                    End If
                End If
            End If
        End If
    Next
End Sub


