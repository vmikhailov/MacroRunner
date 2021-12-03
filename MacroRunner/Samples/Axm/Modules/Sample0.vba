Public C_PriceBaseS As Long

Function ReadDataESCh(path As String, FilesCh As String, SheetCh As String) As Dictionary
    Dim Chw As Workbook
    Dim Chws As Worksheet
	Set Chw = Workbooks.Open(path & "\" & FilesCh)
	Set Chws = Chw.Sheets(SheetCh)
	    
    With Chws
        FirstRow_W = .Range("RowName").row + 2
        LastRow_W = .UsedRange.Rows.Count
        C_ArticleS = .Range("ArticleS_FU").Column
        C_NameS = .Range("Name_New").Column
        C_UnitS = .Range("Size_New").Column
        C_PriceS_EUR = .Range("PriceEssEUR").Column
        C_PriceS_RUB = .Range("PriceEssRub").Column
        C_FactorUnitS = .Range("FactorUnit").Column
        C_PriceS_RUB_FU = .Range("PriceEssRub_FU").Column
        C_ProfitS = .Range("ProfitRRC").Column
        C_PriceBaseS_FU = .Range("RRC").Column
        C_StatS = .Range("StatusS_New").Column
        C_SkladS = .Range("SkladS_New").Column
    End With
      
    Dim SuppData As New Dictionary 'é·˙ﬂ‚ÎÂÌËÂ ÔÂÂÏÂÌÌÓÈ ÚËÔ‡ ÄÒÒÓˆË‡ÚË‚Ì˚È Ï‡ÒÒË‚
    Dim row As ImportDataShippCh
    Dim s As Long
    s = 0
    'óÚÂÌËÂ ‰‡ÌÌ˚ı
    For i = FirstRow_W To LastRow_W
        Set row = New ImportDataShippCh
        If Chws.Cells(i, C_ArticleS) <> "" Then
            With row
                .ArticleS = Chws.Cells(i, C_ArticleS)
                .NameS = Chws.Cells(i, C_NameS)
                .UnitS = Chws.Cells(i, C_UnitS)
                .PriceS_EUR = Chws.Cells(i, C_PriceS_EUR)
                .PriceS_RUB = Chws.Cells(i, C_PriceS_RUB)
                .FactorUnitS = Chws.Cells(i, C_FactorUnitS)
                .PriceS_RUB_FU = Chws.Cells(i, C_PriceS_RUB_FU)
                .PriceBaseS_FU = Chws.Cells(i, C_PriceBaseS_FU)
                .ProfitS = Chws.Cells(i, C_ProfitS)
                .StatS = Chws.Cells(i, C_StatS)
                .SkladS = Chws.Cells(i, C_SkladS)
                
            End With
            Set SuppData(row.ArticleS) = row 'á‡ÔËÒ¸ ‰‡ÌÌ˚ı ÔÓ ÍÎ˛˜Û ‚ Ï‡ÒÒË‚
            s = s + 1
        Else: End If
    Next i
    Set ReadDataESCh = SuppData
    MsgBox (s)
EndSub:
    Chw.Close SaveChanges:=False 
End Function

Private Sub DownloadDataEsCh(FilesName As String, SheetName As String)

    Dim C_ArticleS_AvD As Long
    Dim C_NameS_AvD As Long
    Dim C_UnitS_AvD As Long
    Dim C_PriceS_EUR_AvD As Long
    Dim C_PriceS_RUB_AvD As Long
    Dim C_FactorUnitS_AvD As Long
    Dim C_PriceS_FU_AvD As Long
    Dim C_PriceBaseS_FU_AvD As Long
    Dim C_ProfitS_AvD As Long
    Dim C_StatS_AvD As Long
    Dim C_SkladS_AvD As Long
  
    
    Dim ws As Worksheet
    Set ws = Workbooks(FilesName).Sheets(SheetName)
    
    ws.Activate
    FirstRow = Range("RowName").row + 2
    LastRow = ActiveSheet.UsedRange.Rows.Count
    
    C_ArticleS_AvD = Range("Es_Art_FU").Column
    C_PriceS_EUR_AvD = Range("Es_PriceEUR").Column
    C_PriceS_RUB_AvD = Range("Es_PriceRUB").Column
    C_FactorUnitS_AvD = Range("Es_FactorUnit").Column
    C_PriceS_FU_AvD = Range("Es_PriceRUBFU").Column
    C_PriceBaseS_FU_AvD = Range("Es_RRCFU").Column
    C_ProfitS_AvD = Range("Es_Profit").Column
    C_NameS_AvD = Range("Es_Name").Column
    C_UnitS_AvD = Range("Es_Unit").Column
    C_StatS_AvD = Range("Es_Status").Column
    C_SkladS_AvD = Range("Es_Sklad").Column
    
    'ì‰‡ÎÂÌËÂ ÔÂ‰˚‰Û˘Ëı ‰‡ÌÌ˚ı
    Range(Cells(FirstRow, C_SkladS_AvD), Cells(LastRow, C_SkladS_AvD)).ClearContents
  ' Range(Cells(FirstRow, C_StatS_AvD), Cells(LastRow_W, C_StatS_AvD)).ClearContents
    
  'á‡„ÛÁÍ‡ ‰‡ÌÌ˚ı ËÁ î‡ÈÎ‡ ËÁÏÂÌÂÌËÈ ÔÓÒÚ‡‚˘ËÍ‡
  
    Dim SuppData As Dictionary
    Set SuppData = ReadDataESCh(PathFile, f2, sh2)
    Dim row As ImportDataShippCh

    Dim code As String 'ÄÚËÍÛÎ ÔÓÒÚ‡‚˘ËÍ‡ ‚ ÒÔ‡‚Ó˜ÌËÍÂ AvardaDir
    Summ = 0
    For i = FirstRow To LastRow
        code = ws.Cells(i, C_ArticleS_AvD)
        If SuppData.Exists(code) Then
            Set row = SuppData(code)
        With ws
'ÔÂÂÌÓÒ ‰‡ÌÌ˚ı Ì‡ ÎËÒÚ
           
           .Cells(i, C_PriceS_EUR_AvD).value = row.PriceS_EUR
           .Cells(i, C_PriceS_RUB_AvD).value = row.PriceS_RUB
           .Cells(i, C_FactorUnitS_AvD).value = row.FactorUnitS
           .Cells(i, C_PriceS_FU_AvD).value = row.PriceS_RUB_FU
           .Cells(i, C_PriceBaseS_FU_AvD).value = row.PriceBaseS_FU
           .Cells(i, C_ProfitS_AvD).value = row.ProfitS
           .Cells(i, C_NameS_AvD).value = row.NameS
           .Cells(i, C_UnitS_AvD).value = row.UnitS
           .Cells(i, C_StatS_AvD).value = row.StatS
           .Cells(i, C_SkladS_AvD).value = row.SkladS
           
        End With
        Else: End If
        Summ = Summ + 1
    Next i

End Sub

 Sub DataEsselteCh()
     ' éÚÍÎ˛˜ÂÌËÂ ˝Í‡Ì‡
   Application.ScreenUpdating = False
    
    PathFile = ActiveWorkbook.path
    PathFile = PathFile & "\..\" 'Ì‡ ÛÓ‚ÂÌ¸ ‚˚¯Â
    ' PathFile = "\\SERVER_A\Avarda\Common\àÏÔÓÚ ‰‡ÌÌ˚ı ‚ Ä‚‡‰Û\è‡ÈÒ-ÎËÒÚ˚ ÔÓÒÚ‡‚˘ËÍÓ‚\"
   
    f1 = "AvardaDir.xls"
    sh1 = "ProductAvarda"
    f2 = "àÏÔÓÚ\Esselte_àÁÏÂÌÂÌËﬂ.xls"
    sh2 = "Ç˚·ÓÍ‡ ËÁ Ô‡ÈÒ‡"
    
     datat1 = Now()
     Call DownloadDataEsCh(f1, sh1)
	 Dim aa as Integer
     datat2 = Now()
     datat2 = datat2 - datat1

	 'comment line
     MsgBox (datat2 & "ÒÂÍ. á‡„ÛÁÍ‡ ‰‡ÌÌ˚ı  - " & Summ)
     'comment line

   Application.ScreenUpdating = True
 End Sub

