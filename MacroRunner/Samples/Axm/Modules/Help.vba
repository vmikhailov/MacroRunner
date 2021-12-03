Sub ArtKomusSearch()
'
' å‡ÍÓÒ1 å‡ÍÓÒ
' å‡ÍÓÒ Á‡ÔËÒ‡Ì 09.10.2014 (mlm)
    Dim C_art As Long
    Dim CK_art As Long
    Dim n As Integer
    
    
    C_art = Range("ArticleAv_New").Column
    ActiveCell.Select
    CK_art = ActiveCell.Column
    n = CK_art - C_art
    ActiveCell.Select
    ActiveCell.FormulaR1C1 = "=RIGHT(RC[-29],LEN(RC[-29])-FIND(""-"",RC[-29],1))"
   'ActiveCell.FormulaR1C1 = "=RIGHT(Cells(R_art, C_art),LEN(Cells(R_art, C_art))-FIND(""-"",Cells(R_art, C_art),1))"
  ' ActiveCell.Value = Right(Cells(R_art, C_art), (Len(Cells(R_art, C_art)) - Find("-", Cells(R_art, C_art), 1))) 'Find
End Sub
Sub ArtKomusSearch_6()
'
' å‡ÍÓÒ1 å‡ÍÓÒ
' å‡ÍÓÒ Á‡ÔËÒ‡Ì 09.10.2014 (mlm)
'
    Dim C_art As Long
    Dim R_art As Long
    
    C_art = Range("ArticleAv_New").Column
    ActiveCell.Select
    R_art = ActiveCell.row
   ' ActiveCell.FormulaR1C1 = "=RIGHT(RC[-27],6)"
     ActiveCell.value = Right(Cells(R_art, C_art), 6)
   
End Sub
Sub ArtKomusSearch_5()
'
    Dim C_art As Long
    Dim R_art As Long
    
    C_art = Range("ArticleAv_New").Column
    ActiveCell.Select
    R_art = ActiveCell.row
   ' ActiveCell.FormulaR1C1 = "=RIGHT(RC[-27],5)"
     ActiveCell.value = Right(Cells(R_art, C_art), 5)
End Sub

Sub AutoF()
   
    Dim FirstCol As Long, LastCol As Long
    Dim FirstRow As Long, LastRow As Long
    
    FirstRow = 2 'ActiveCell.row
    LastRow = ActiveSheet.UsedRange.Rows.Count
    FirstCol = 1
    LastCol = ActiveSheet.UsedRange.Columns.Count
    
    With ActiveSheet
         .AutoFilterMode = False
         .Range(Cells(FirstRow, FirstCol), Cells(LastRow, LastCol)).AutoFilter
    End With
End Sub
Sub GroupingAv()
 Dim i As Long
 Dim NCol As Long
 Dim FirstRow, LastRow As Long
 
    NCol = Range("Unit").Column
    'LastRow = ActiveSheet.UsedRange.Rows.Count
    LastRow = 7154
    FirstRow = 4
    ActiveSheet.Outline.ShowLevels RowLevels:=2
    Cells.Select
    Application.CutCopyMode = False
    Selection.Rows.Ungroup
    Range(Cells(FirstRow, 1), Cells(LastRow, 256)).Rows.Group
    
    For i = FirstRow To LastRow
       If Cells(i, NCol).Value = "aaa" Then

            Range(Cells(i, 1), Cells(i, 256)).Rows.Ungroup
       Else
            If Cells(i, NCol).value = "ÉêìèèÄ" Then
               Range(Cells(i, 1), Cells(i, 256)).Rows.Ungroup
            Else
                If Cells(i, NCol).value = "èéÑÉê" Then
                    Range(Cells(i, 1), Cells(i, 256)).Rows.Ungroup
                Else: End If
            End If
        End If
    Next i
    Range("A1").Select
    ActiveSheet.Outline.ShowLevels RowLevels:=1
End Sub
