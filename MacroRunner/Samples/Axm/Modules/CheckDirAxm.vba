Function ReadDirectoryAxm(path As String, Filename As String, SheetName As String) As Dictionary
    Dim w As Workbook
    Dim ws As Worksheet
    Set w = Workbooks.Open(path & "\" & Filename)
    Set ws = w.Sheets(SheetName)
    
    Dim ColArt As Long
    
    FirstRow = Range("ÌÂ_Û‰‡ÎﬂÚ¸").row + 1
    LastRow = ws.UsedRange.Rows.Count
    ColCode = Range("KanzCodeAv").Column
    ColArt = Range("KanzArticle").Column
    Dim SuppData As New Dictionary 'é·˙ﬂ‚ÎÂÌËÂ ÔÂÂÏÂÌÌÓÈ ÚËÔ‡ ÄÒÒÓˆË‡ÚË‚Ì˚È Ï‡ÒÒË‚
    Dim row As ImportDataAxm
    Dim s As Long
    s = 0
    For i = FirstRow + 1 To LastRow
        Set row = New ImportDataAxm
        row.CodeKanz = ws.Cells(i, ColCode)
        row.ArticleKanz = ws.Cells(i, ColArt)
        Set SuppData(row.CodeKanz) = row 'á‡ÔËÒ¸ ‰‡ÌÌ˚ı ÔÓ ÍÎ˛˜Û ‚ Ï‡ÒÒË‚
        s = s + 1
    Next i
    Set ReadDirectoryAxm = SuppData
    MsgBox (s)
EndSub:
    w.Close SaveChanges:=False
End Function
Private Sub DownloadArticleAxm(FilesCh As String, SheetCh As String)

    Dim SuppData As Dictionary
    Set SuppData = ReadDirectoryAxm(PathFile, f2, sh2)
    Dim row As ImportDataAxm
    Dim Chws As Worksheet
    Set Chws = Workbooks(FilesCh).Sheets(SheetCh)
    
    FirstRow_W = Range("RowName").row + 1
    LastRow_W = Chws.UsedRange.Rows.Count
    ColumnCodeKanz = Range("CodeAv").Column
    Dim ColArtAxm As Long
    Dim code As String 'ÄÚËÍÛÎ ‚ Ô‡ÈÒÂ ËÁÏÂÌÂÌËÈ
    
    ColArtAxm = Range("ArticleDirAxm").Column
    Summ = 0
    SummPr = 0
    For i = FirstRow_W To LastRow_W
        code = Chws.Cells(i, ColumnCodeKanz)
         If SuppData.Exists(code) Then
            Set row = SuppData(code)
     
            Chws.Cells(i, ColArtAxm).value = row.ArticleKanz
         Else
            
         End If
    Next i
   Summ = SummPr + SummNew
  
End Sub

 Sub LoadArtAxm()
     ' éÚÍÎ˛˜ÂÌËÂ ˝Í‡Ì‡
   Application.ScreenUpdating = False
    
    PathFile = ActiveWorkbook.path
    PathFile = PathFile & "\..\..\" ' Ì‡ ÛÓ‚ÂÌ¸ ‚˚¯Â
    
    f2 = "ëÔ‡‚Ó˜ÌËÍíÓ‚‡Ó‚New.xls"
    sh2 = "ëÔËÒÓÍíÓ‚‡Ó‚"
    f1 = "AvardaDir.xls"
    sh1 = "ProductAvarda"
       
    Call DownloadArticleAxm(f1, sh1)
   MsgBox ("á‡„ÛÁËÎË =))")
         ' ÇÍÎ˛˜ÂÌËÂ ˝Í‡Ì‡
   Application.ScreenUpdating = True
 End Sub



